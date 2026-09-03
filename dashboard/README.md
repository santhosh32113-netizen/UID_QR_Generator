# Interactive panel data

Run this from the project root whenever the source workbook changes:

```bash
./venv/bin/python tools/create_dashboard_data.py
```

To use the Add asset screen and write entries back to Excel, run the persistence-enabled server:

```bash
./venv/bin/python tools/dashboard_server.py
```

Then open `http://127.0.0.1:8765/index.html`.

This creates:

- `dashboard/fleet_register.csv` for Power BI
- `dashboard/data.js` for the local interactive panel

## Power BI setup

1. Select **Get data > Text/CSV** and choose `dashboard/fleet_register.csv`.
2. Use `Serv`, `Type`, `Brigade`, `Division`, `Corps`, `Command`, `Unit`, `OEM`, `Payload`, and `Guidance` as slicers.
3. Use `Drone ID` as the asset-level table field.
4. Add measures for row count, serviceable count, and total `Cost (in Thousands)`.

The CSV intentionally contains the complete operational record. Keep the restricted `Drone ID | KUIN-G` master register out of the report dataset.

Drone ID is not editable. It is generated from the first 2 characters of Command and Corps; first 3 of Division, Drone Name, Type, and Form Factor; and first 4 of Brigade and Unit. A numeric suffix is added when that prescribed abbreviated hierarchy repeats, preserving primary-key uniqueness.

New dropdown values are resolved to their companion text input before saving; the literal `__new__` marker is never written to Excel.

Default local credentials are `admin` / `admin` and `user` / `user`. Change them from the Password control after signing in. Users can add records; only Admin can view the full register or delete records.