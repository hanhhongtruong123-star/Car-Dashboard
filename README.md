# Car Service Dashboard

KPI dashboard for the car service (Vegas Car fleet, with an optional Grab module) for OD.
The whole dashboard is **one HTML file**: open it in any browser, no server needed.

> **Demo data only.** This repository contains *fake* sample data (`data/sample/`).
> Real Excel files (guest IDs, staff names, costs) must never be committed — put them in `data/private/`, which git ignores.

## Modules
Overview · Customers · VG Cars · Drivers · Car Cost · Grab Cost (hidden by default, switched on in **Admin**).
Filters: period (day / month / quarter / year), compare (previous period / same period last year), trip type (all / guest / internal).

## Build

Requirements: Python 3 with `openpyxl`, Node.js (only to download Chart.js and SheetJS once).

```bash
npm install
pip install openpyxl
python3 build.py                              # sample data  -> dist/index.html
python3 build.py --data data/private          # your real files (stay local)
python3 build.py --cdn                        # smaller file; loads libraries from cdnjs when opened
```

Every sheet of every `.xlsx` in `--data` becomes the built-in data. New files can also be dropped on the page (**Admin → upload**).

## Accepted Excel layouts

| File | Sheet / columns |
|---|---|
| VG Car trips | `GuestID, Type, Customer Tier` (or `Level`), `Nationality` (or `Nation`), `Date, Depart, Arrival, Occupancy, Distance, Arrangby, EmployeeId, Brand, Destination, CITY`; optional `Confirm` (guard name), `Host Team` |
| Driver names | `EmployeeId, Name` (or `Driver Name`) |
| Car cost | `Date, Car Brand, Cost Type, Amount (Usd)` (or `Date, Vehicle, Category, Amount (USD)`) |
| Grab | sheet name contains `Grab`; trip columns plus `Grab $` and `Hour` |
| Driver roster | `Date, EmployeeId, Start, End, Break, Status` — see `templates/Driver_Roster_Template.xlsx` |

## Placeholders still in use
- **Host team** split is simulated (each guest trip gets a random but stable team) until a `Host Team` column is supplied.
- **Guard** = `Arrangby` (security team); the name comes from the `Confirm` column.

## Project layout
```
src/template.html   page layout and styles (light + dark)
src/app.js          parsing, KPIs and charts
build.py            builds dist/index.html
data/sample/        fake demo data (safe to publish)
templates/          Excel template for the driver roster
.github/workflows/  publishes the sample build to GitHub Pages
```

## Demo link (GitHub Pages)
Repository **Settings → Pages → Source: GitHub Actions**. Each push to `main` rebuilds the demo from `data/sample/` and publishes it at `https://<owner>.github.io/<repo>/`.
