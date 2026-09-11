# pqdoc

Power BI Dataflow (Gen1 / Power Query) export JSON → Markdown documentation tool.

Reads one or more dataflow export JSON files (exported from Power BI / Fabric
dataflows) and produces a per-dataflow Markdown document covering:

1. **All data sources** — connector + connection string / upstream dataflow
   entity, and which queries use them
2. **Transformation steps** — every step, named, in order, with the PQ function
   each one calls
3. **Output tables** — entities with columns (name + type), refresh location
   and load status
4. **Lineage** — each output table traced back to its ultimate source(s)

## Usage

```bash
python3 pqdoc.py file1.json [file2.json ...] [-o outdir]
```

- Writes `<name>.pqdoc.md` per dataflow into `outdir` (defaults to the
  directory of the JSON files).
- Prints a one-screen summary per dataflow (sources, outputs, lineage) to
  stdout.

Pure Python stdlib only — no dependencies. Python 3.8+.

## Examples

```bash
# document a single export, output alongside the JSON
python3 pqdoc.py parkrun-participation-2026.json

# batch several, dump markdown into ./out
python3 pqdoc.py a.json b.json c.json -o out
```

## Sample output

See `out/` for three real generated documents
(`parkrun-participation-2026`, `parkrun-volunteers-history-2026`,
`sharepoint-people-2026`).
