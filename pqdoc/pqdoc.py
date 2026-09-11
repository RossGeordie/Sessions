#!/usr/bin/env python3
"""
pqdoc — Power BI Dataflow (Gen1 / Power Query) export JSON documentation tool.

Reads one or more dataflow export JSON files (exported from Power BI /
Fabric dataflows) and produces a per-dataflow Markdown document covering:

  1. ALL data sources        — connector + connection string / upstream
                               dataflow entity, and which queries use them
  2. Transformation steps    — every step, named, in order, with the PQ
                               function each one calls
  3. Output tables           — entities with columns (name + type),
                               refresh location and load status
  4. Lineage                 — each output table traced back to its
                               ultimate source(s)

Usage:
  python3 pqdoc.py file1.json [file2.json ...] [-o outdir]
"""

import argparse
import json
import os
import re
import sys
from collections import OrderedDict

# ---------------------------------------------------------------------------
# Connector / source detection
# ---------------------------------------------------------------------------
_CONNECTORS = [
    ("SharePoint.Files",    re.compile(r'SharePoint\.Files\(\s*"[^"]*"')),
    ("SharePoint.Tables",   re.compile(r'SharePoint\.Tables\(\s*"[^"]*"')),
    ("SharePoint.List",     re.compile(r'SharePoint\.List\(\s*"[^"]*"')),
    ("Web.Contents",        re.compile(r'Web\.Contents\(\s*"[^"]*"')),
    ("Web.Page",            re.compile(r'Web\.Page\(\s*"[^"]*"')),
    ("Excel.OnDataRange",   re.compile(r'Excel\.OnDataRange\(\s*"[^"]*"')),
    ("Excel.Workbook",      re.compile(r'\bExcel\.Workbook\b')),
    ("OData.Feed",          re.compile(r'OData\.Feed\(\s*"[^"]*"')),
    ("Sql.Database",        re.compile(r'Sql\.Database\(\s*"[^"]*"\s*,\s*"[^"]*"')),
    ("AzureSql.Database",   re.compile(r'AzureSql\.Database\(\s*"[^"]*"\s*,\s*"[^"]*"')),
    ("PostgreSQL.Database", re.compile(r'PostgreSQL\.Database\(\s*"[^"]*"')),
    ("MySQL.Database",      re.compile(r'MySql\.Database\(\s*"[^"]*"')),
    ("Oracle.Database",     re.compile(r'Oracle\.Database\(\s*"[^"]*"')),
    ("Odbc",                re.compile(r'Odbc\(\s*"[^"]*"')),
    ("Csv.Document",        re.compile(r'\bCsv\.Document\b')),
    ("Json.Document",       re.compile(r'\bJson\.Document\b')),
    ("Xml.Document",        re.compile(r'\bXml\.Document\b')),
    ("Parquet.Document",    re.compile(r'Parquet\.Document\b')),
    ("GoogleSheets",        re.compile(r'GoogleSheets\b')),
    ("Kusto",               re.compile(r'Kusto\.FromCluster\(\s*"[^"]*"')),
    ("PowerBI.Dataflow",    re.compile(r'PowerBI\.\w+\b')),
    ("PowerPlatform.Dataflows", re.compile(r'PowerPlatform\.Dataflows\b')),
    ("AzureBlob",           re.compile(r'AzureBlob\.\w+\(\s*"')),
    ("OneDrive",            re.compile(r'OneDrive\.Folders?\(\s*"')),
    ("Folder",              re.compile(r'\bFolder\.\w+\(\s*"')),
    ("File.Contents",       re.compile(r'File\.Contents\(\s*"[^"]*"')),
]

# document-of-binary: read [Content] from an upstream query, not a direct conn
_INLINE_DOCS = {"Csv.Document", "Json.Document", "Xml.Document",
                "Excel.Workbook", "Parquet.Document"}


def find_connectors(expr):
    """Return list of (connector_label, connection_string_or_None) found in expr."""
    found = []
    for label, rx in _CONNECTORS:
        m = rx.search(expr)
        if m:
            q = re.search(r'"([^"]*)"', m.group(0))
            found.append((label, q.group(1) if q and q.lastindex else None))
    return found


def referenced_queries(expr, known_names):
    """Identify other local shared queries referenced inside an expression."""
    refs = set()
    for name in known_names:
        if re.search(r'(?<![\w."])' + re.escape(name) + r'(?![\w])', expr):
            refs.add(name)
    return refs

# identifiers that should not be mistaken for upstream dataflow entities
_PQ_BUILTINS = {
    "Table", "Text", "List", "Excel", "SharePoint", "Json", "Csv", "Xml", "Web",
    "Type", "Date", "DateTime", "Duration", "Number", "Binary", "Time", "Record",
    "Value", "Function", "Section", "Int64", "Int32", "Int16", "Int8", "Splitter",
    "JoinKind", "QuoteStyle", "Replacer", "ValueFieldType", "Order", "TableShape",
    "null", "each", "in", "let", "type", "meta", "as", "if", "then", "else",
    "for", "and", "or", "true", "false", "null", "Source", "Error", "ExtraValues",
    "Value", "Binary", "Records", "Columns", "Data", "Item", "Content",
}


def first_function(expr):
    """Best-effort: the PQ function a step's expression primarily calls."""
    expr = expr.strip()
    m = re.match(r'(?:#\w+|[\w.]+)\s*\(', expr)
    if m:
        return m.group(0).rstrip("(").strip()
    m = re.search(r'=\s*(?:each\s+)?#"?([\w.]+)\s*\(', expr)
    if m:
        return m.group(1)
    return ""

# ---------------------------------------------------------------------------
# Query model
# ---------------------------------------------------------------------------
class Query:
    def __init__(self, name):
        self.name = name
        self.steps = []          # [(step_name, expression)]
        self.expr = ""           # full let..in body
        self.is_parameter = False

    @property
    def source_step(self):
        for nm, ex in self.steps:
            if nm.lower() == "source":
                return ex
        return self.steps[0][1] if self.steps else ""


_STEP_NAME_RX = re.compile(r'^\s*(?:#?"([^"]+)"|#?([A-Za-z_][\w ]*))\s*=\s*(.*)$')


def parse_document(doc):
    queries = OrderedDict()
    headers = list(re.finditer(r'(?m)^\s*shared\s+("([^"\n]+)"|([A-Za-z_][\w]*))\s*=\s*let\b', doc))
    for i, h in enumerate(headers):
        name = (h.group(2) or h.group(3)).strip()
        start = h.end()
        end = headers[i + 1].start() if i + 1 < len(headers) else len(doc)
        body = doc[start:end]

        q = Query(name)
        if "IsParameterQuery" in body:
            q.is_parameter = True

        in_match = None
        for m in re.finditer(r'\r?\n[ \t]*\bin\b(?![\w])', body):
            in_match = m
        steps_region = body[:in_match.start()] if in_match else body

        for line in steps_region.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("//"):
                continue
            m = _STEP_NAME_RX.match(line)
            if m:
                sname = (m.group(1) or m.group(2)).strip()
                q.steps.append((sname, m.group(3).strip().rstrip(",").strip()))
            elif q.steps:
                nm, ex = q.steps[-1]
                q.steps[-1] = (nm, ex.rstrip(",") + " " + stripped.rstrip(","))
        q.expr = body.strip()
        queries[name] = q
    return queries

# ---------------------------------------------------------------------------
# Source classification & lineage
# ---------------------------------------------------------------------------
def dep_graph(all_queries):
    """{query_name: set(local names it references anywhere)}"""
    g = {}
    for name, q in all_queries.items():
        others = set(all_queries) - {name}
        g[name] = referenced_queries(q.expr, others)
    return g


def leaves_of(start_name, g, all_queries):
    """Walk the dependency graph DOWN to leaves.
    Returns (local_leaf_list, external_entity_names)."""
    local = set(all_queries)
    leaves, exts, seen = [], set(), set()
    stack = [start_name]
    while stack:
        cur = stack.pop()
        if cur in seen:
            continue
        seen.add(cur)
        q = all_queries.get(cur)
        if q is None:
            exts.add(cur)
            continue
        deps = g.get(cur, set())
        if not deps:
            if getattr(q, "is_parameter", False):
                leaves.append(("parameter", f"sample file — {cur}"))
            else:
                conns = find_connectors(q.source_step)
                if conns:
                    label, cs = conns[0]
                    if label in _INLINE_DOCS:
                        leaves.append(("embedded-binary", f"[Content] binary — {cur}"))
                    elif label == "PowerPlatform.Dataflows":
                        me = re.search(r'\[entity\s*=\s*"([^"]+)"', q.expr)
                        mi = re.search(r'\[dataflowId\s*=\s*"([^"]+)"', q.expr)
                        ent = me.group(1) if me else "?"
                        leaves.append(("upstream-dataflow", f"{ent} (dataflow {mi.group(1)[:8] if mi else '?'}…)" ))
                    else:
                        leaves.append((label, cs or "n/a"))
                else:
                    exts.add(cur)
            continue
        stack.extend(d for d in deps if d not in seen)
    return leaves, exts


def classify_source(query, all_queries, upstream):
    """(kind, detail) for the query's source step."""
    if getattr(query, "is_parameter", False):
        return "parameter", "binary/parameter sample"
    src = query.source_step
    conns = find_connectors(src)
    if conns:
        label, cs = conns[0]
        if label in _INLINE_DOCS:
            return "embedded-binary", "document of [Content] from upstream query"
        if label == "PowerPlatform.Dataflows":
            m_ent = re.search(r'\[entity\s*=\s*"([^"]+)"', query.expr)
            m_id = re.search(r'\[dataflowId\s*=\s*"([^"]+)"', query.expr)
            ent = m_ent.group(1) if m_ent else "?"
            df = m_id.group(1) if m_id else "?"
            return "upstream-dataflow", f"{ent}  (dataflow {df[:8]}…)"
        return label, cs or "n/a"
    refs = referenced_queries(src, set(all_queries) - {query.name})
    if refs:
        return "shared-query", ", ".join(sorted(refs))
    # otherwise: bare identifier = upstream dataflow entity (or external ref)
    ext = None
    for m in re.finditer(r'(?<![\w."])([A-Za-z_][\w]*)(?![\w(])', src):
        tok = m.group(1)
        if tok in _PQ_BUILTINS or tok in all_queries:
            continue
        ext = tok
        break
    if ext:
        return ("upstream-dataflow" if upstream else "external-entity"), ext
    return "unknown", None

# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------
def slugify(name):
    return re.sub(r'[^A-Za-z0-9]+', '-', name).strip('-').lower()


def build_report(path):
    with open(path) as f:
        d = json.load(f)
    mashup = d.get("pbi:mashup", {})
    doc = mashup.get("document", "")
    qmeta = mashup.get("queriesMetadata", {})
    conn_overrides = mashup.get("connectionOverrides", [])
    entities = d.get("entities", [])

    name = d.get("name", os.path.basename(path))
    queries = parse_document(doc)
    g = dep_graph(queries)
    upstream = any("PowerPlatformDataflows" in json.dumps(c) for c in conn_overrides)

    # --- sources (deduped by connector+connection or entity) ---
    sources = OrderedDict()
    query_sources = OrderedDict()
    for qname in queries:
        kind, detail = classify_source(queries[qname], queries, upstream)
        if kind in ("shared-query", "unknown"):
            continue
        key = (kind, detail or "n/a")
        sources.setdefault(key, {"kind": kind, "detail": detail or "n/a", "queries": []})
        sources[key]["queries"].append(qname)
        query_sources[qname] = (kind, detail)

    # --- steps ---
    steps_report = OrderedDict()
    for qname, q in queries.items():
        steps_report[qname] = {
            "steps": [{"name": s[0], "fn": first_function(s[1])} for s in q.steps],
            "load_enabled": qmeta.get(qname, {}).get("loadEnabled"),
            "source": query_sources.get(qname),
        }

    # --- outputs ---
    outputs = []
    for e in entities:
        outputs.append({
            "name": e.get("name"),
            "description": e.get("description", ""),
            "columns": [(a.get("name"), a.get("dataType", "")) for a in e.get("attributes", [])],
            "refresh_location": e.get("pbi:refreshPolicy", {}).get("location", ""),
        })

    # --- lineage: output -> leaves ---
    out_lineage = {}
    for o in outputs:
        oname = o["name"]
        if oname in queries:
            leaves, exts = leaves_of(oname, g, queries)
        else:
            q = queries.get(oname)
            leaves = []
            exts = set()
        out_lineage[oname] = {"leaves": leaves, "exts": sorted(exts)}

    return {
        "file": path,
        "name": name,
        "description": d.get("description", ""),
        "modified": d.get("modifiedTime", ""),
        "outfmt": d.get("ppdf:outputFileFormat", "n/a"),
        "sources": sources,
        "steps_report": steps_report,
        "outputs": outputs,
        "out_lineage": out_lineage,
        "conn_overrides": conn_overrides,
        "query_count": len(queries),
        "upstream": upstream,
    }


def render_markdown(r):
    L = []
    L.append(f"# {r['name']}")
    L.append("")
    if r["description"]:
        L.append(f"> {r['description']}")
        L.append("")
    L.append(f"- **File:** `{r['file']}`")
    if r["modified"]:
        L.append(f"- **Modified:** {r['modified']}")
    L.append(f"- **Output format:** {r['outfmt']}")
    L.append(f"- **Shared queries:** {r['query_count']}")
    if r["upstream"]:
        L.append("- **Consumes upstream dataflows** (PowerPlatformDataflows connection)")
    L.append("")

    # sources
    L.append("## 1. Data Sources")
    L.append("")
    if not r["sources"]:
        L.append("_No direct external connectors found._")
        L.append("")
    else:
        L.append("| # | Type | Connection / Entity | Used by queries |")
        L.append("|---|------|---------------------|-----------------|")
        for i, (key, s) in enumerate(r["sources"].items(), 1):
            dq = ", ".join(f"`{q}`" for q in s["queries"])
            L.append(f"| {i} | **{s['kind']}** | `{s['detail']}` | {dq} |")
        L.append("")
        kinds = {k for k in r["sources"]}
        if "parameter" in kinds:
            L.append("> `parameter` = a sample-file binary parameter (PQ parameter query).")
            L.append("")
        if "upstream-dataflow" in kinds or "external-entity" in kinds:
            L.append("> Upstream-entity rows mean this flow reads a table published by ANOTHER dataflow — follow that name to the upstream flow that publishes it.")
            L.append("")

    # steps
    L.append("## 2. Transformation Steps (per query)")
    L.append("")
    for qname, info in r["steps_report"].items():
        src = info["source"]
        if src:
            src_txt = f"source: `{src[0]}`" + (f" — `{src[1]}`" if src[1] else "")
        else:
            src_txt = "source: (references another query)"
        load = info["load_enabled"]
        load_txt = {True: "→ loads to output", False: "**not loaded** (disabled)", None: "no load flag"}.get(load, "")
        L.append(f"### `{qname}`")
        L.append("")
        L.append(f"_{src_txt} · {load_txt}_")
        L.append("")
        if not info["steps"]:
            L.append("_no steps found_")
            L.append("")
            continue
        L.append("| # | Step name | Function |")
        L.append("|---|-----------|----------|")
        for i, s in enumerate(info["steps"], 1):
            fn = f"`{s['fn']}`" if s["fn"] else "—"
            L.append(f"| {i} | {s['name']} | {fn} |")
        L.append("")

    # outputs
    L.append("## 3. Output Tables")
    L.append("")
    if not r["outputs"]:
        L.append("_No output entities declared._")
        L.append("")
    for o in r["outputs"]:
        L.append(f"### `{o['name']}`")
        L.append("")
        if o["description"]:
            L.append(f"> {o['description']}")
            L.append("")
        if o["refresh_location"]:
            L.append(f"- **Refresh location:** `{o['refresh_location']}`")
        L.append(f"- **Columns: {len(o['columns'])}**")
        L.append("")
        L.append("| Column | Type |")
        L.append("|--------|------|")
        for cname, ctype in o["columns"]:
            L.append(f"| {cname} | {ctype} |")
        L.append("")

    # lineage
    L.append("## 4. Lineage — output ← sources")
    L.append("")
    for o in r["outputs"]:
        lin = r["out_lineage"].get(o["name"], {})
        parts = [f"`{o['name']}`"]
        kind_map = {
            "parameter": "📁 sample-file param",
            "embedded-binary": "📄 binary doc",
            "upstream-dataflow": "⬆ upstream dataflow",
            "external-entity": "⬆ external entity",
        }
        for leaf, detail in lin.get("leaves", []):
            marker = kind_map.get(leaf, leaf)
            parts.append(f"{marker}: `{detail}`")
        for ext in lin.get("exts", []):
            parts.append(f"⬆ upstream dataflow: `{ext}`")
        L.append(" ← ".join(reversed(parts)) + "  ")
    L.append("")

    # conn overrides
    if r["conn_overrides"]:
        L.append("## 5. Connection Overrides")
        L.append("")
        for c in r["conn_overrides"]:
            L.append(f"- {json.dumps(c)}")
        L.append("")

    L.append("---")
    L.append(f"_Generated by pqdoc from `{r['file']}`_")
    L.append("")
    return "\n".join(L)


def render_summary(r):
    lines = [f"== {r['name']} ({r['file']}) =="]
    lines.append(f"  sources {len(r['sources'])} · queries {r['query_count']} · outputs {len(r['outputs'])}")
    for key, s in r["sources"].items():
        lines.append(f"    SRC  {s['kind']:<18} {s['detail']}   ← {', '.join(s['queries'])}")
    for o in r["outputs"]:
        lin = r["out_lineage"].get(o["name"], {})
        parts = []
        for leaf, detail in lin.get("leaves", []):
            parts.append(f"{leaf}:{detail}")
        for ext in lin.get("exts", []):
            parts.append(f"UPSTREAM:{ext}")
        lines.append(f"    OUT  {o['name']} ({len(o['columns'])} cols)   ←  {' + '.join(parts) or '?'}")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="Document Power BI dataflow export JSON files")
    ap.add_argument("files", nargs="+")
    ap.add_argument("-o", "--outdir", default=None)
    args = ap.parse_args()

    ok, reports = [], []
    for f in args.files:
        if not os.path.isfile(f):
            print(f"!! missing: {f}", file=sys.stderr)
            continue
        try:
            reports.append(build_report(f))
            ok.append(f)
        except Exception as e:
            print(f"!! failed: {f}: {e!r}", file=sys.stderr)

    if ok:
        out_dir = args.outdir or os.path.dirname(os.path.abspath(ok[0]))
        os.makedirs(out_dir, exist_ok=True)
    for r in reports:
        md = os.path.join(out_dir, slugify(r["name"]) + ".pqdoc.md")
        with open(md, "w") as fh:
            fh.write(render_markdown(r))
        print(f"wrote {md}")
    for r in reports:
        print()
        print(render_summary(r))
        print()


if __name__ == "__main__":
    main()
