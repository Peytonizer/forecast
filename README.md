# forecast

Fifteen-year maintenance forecast workbooks for strata. List each task and how often it falls
due; the workbook projects it across 15 years and shows any 3–5 year window, so you can see what
a maintenance contract tender needs to cover. Part of the [strata](https://strata.noradz.io/)
family. Published at <https://forecast.noradz.io/>.

It forecasts tasks, not costs.

## Workbooks

| Trade | File | Status |
| --- | --- | --- |
| Fire | [`files/fire-maintenance-forecast.xlsx`](files/fire-maintenance-forecast.xlsx) | available |
| Plumbing | — | planned |

Each workbook has the same sheets:

- **Register** — one row per task per building: frequency, first due year, quantity and
  notes. The 15-year grid fills itself.
- **Window** — choose a start year, 3, 4 or 5 years, and optionally one building; it lists every
  task due in that window, year by year.
- **Summary** — services due by building and year, with a chart.
- **Settings** — start year, building names, frequencies.

Each workbook covers four buildings being managed. Building names and example tasks are placeholders — overwrite them.

## Building

The workbooks are generated, so a new trade is a new entry in `tools/build_xlsx.py`:

```
python3 tools/build_xlsx.py fire     # needs openpyxl
node scripts/build.mjs               # builds the site into dist/
```

The site is a single static page with the strata palette (`vendor/strata-kit`, a submodule)
inlined at build time; it makes no network requests. Clone with `--recurse-submodules`.

## Licence

MIT.
