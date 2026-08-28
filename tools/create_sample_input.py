import hashlib
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation

OUTPUT = Path("input") / "sample.xlsx"

HEADERS = [
    "Ser No",
    "Drone ID",
    "Drone Name",
    "Type",
    "Form Factor",
    "OEM",
    "Range",
    "Weight (KG)",
    "Endurance (min)",
    "Payload",
    "Payload Weight",
    "Payload Description",
    "Guidance",
    "Anti Ew",
    "C2 Link Frequency",
    "Proc Fund",
    "Serv",
    "Cost (in Thousands)",
    "Image",
    "Unit",
    "Brigade",
    "Division",
    "Corps",
    "Command",
]

TYPES = ["Trg", "Svl (SR)", "Svl (MR)", "FPV", "Kamikaze", "Lgs"]
FORM_FACTORS = ["QC", "HC", "Fixed Wg", "Swarm"]
OEM_NAMES = ["OEM Alpha", "OEM Bravo", "OEM Delta", "OEM Nova", "OEM Zenith"]
ENDURANCE = ["20 min", "35 min", "1 hr", "2 hr", "8 hr"]
PAYLOADS = ["Day Night Cam", "Thermal Cam", "Mapping Sensor", "Training Camera", "None"]
GUIDANCE = ["Comd", "GNSS (GPS)", "Spool", "Tethered drone", "Not Specified"]
ANTI_EW = ["RTH", "Anti-jam", "None", "Not Specified"]
LINK_FREQ = ["2.4 GHz", "5.8 GHz", "900 MHz"]
PROC_FUND = ["ATG", "Regtl", "Unit", "Central"]
SERVICE_STATUS = ["Svc", "Unsvc", "Decommissioned"]
UNITS = [f"Unit {index}" for index in range(1, 13)]
BRIGADES = [f"Brigade {index}" for index in range(1, 7)]
DIVISIONS = ["Division Alpha", "Division Bravo"]
CORPS = ["Corps North", "Corps South"]
COMMANDS = ["Command East", "Command West"]


def make_drone_id(row: dict) -> str:
    fields = "|".join(str(row[header]) for header in HEADERS if header != "Drone ID")
    digest = hashlib.sha256(fields.encode("utf-8")).hexdigest().upper()
    return f"DUMMY-{digest[:4]}-{digest[4:8]}-{digest[8:12]}"


def make_rows(count: int = 120) -> list[list[object]]:
    rows = []
    for serial in range(1, count + 1):
        unit = UNITS[(serial - 1) % len(UNITS)]
        formation = FORMATIONS[(serial - 1) % len(FORMATIONS)]
        row = {
            "Ser No": serial,
            "Drone Name": f"Demo Drone {serial:03d}",
            "Type": TYPES[(serial - 1) % len(TYPES)],
            "Form Factor": FORM_FACTORS[(serial - 1) % len(FORM_FACTORS)],
            "OEM": OEM_NAMES[(serial - 1) % len(OEM_NAMES)],
            "Range": ["< 5 km", "5-10 km", "10-30 km", "31-100 km", "> 100 km"][(serial - 1) % 5],
            "Weight (KG)": 1 + (serial % 8),
            "Endurance (min)": 20 + (serial % 8) * 10,
            "Payload": PAYLOADS[(serial - 1) % len(PAYLOADS)],
            "Payload Weight": round((serial % 20) / 10, 1),
            "Payload Description": f"Synthetic payload profile {((serial - 1) % 5) + 1}",
            "Guidance": GUIDANCE[(serial - 1) % len(GUIDANCE)],
            "Anti Ew": ANTI_EW[(serial - 1) % len(ANTI_EW)],
            "C2 Link Frequency": LINK_FREQ[(serial - 1) % len(LINK_FREQ)],
            "Proc Fund": PROC_FUND[(serial - 1) % len(PROC_FUND)],
            "Serv": SERVICE_STATUS[0] if serial % 9 else SERVICE_STATUS[1],
            "Cost (in Thousands)": 250 + ((serial * 37) % 4750),
            "Image": "Dummy image placeholder",
            "Unit": unit,
            "Brigade": BRIGADES[(serial - 1) % len(BRIGADES)],
            "Division": DIVISIONS[(serial - 1) % len(DIVISIONS)],
            "Corps": CORPS[(serial - 1) % len(CORPS)],
            "Command": COMMANDS[(serial - 1) % len(COMMANDS)],
        }
        row["Drone ID"] = make_drone_id(row)
        rows.append([row[header] for header in HEADERS])
    return rows


def add_dropdown(ws, formula: str, cell_range: str) -> None:
    validation = DataValidation(type="list", formula1=f'"{formula}"', allow_blank=False)
    ws.add_data_validation(validation)
    validation.add(cell_range)

wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

ws["A1"] = "Dummy Drone Inventory - Synthetic Data"
for column, header in enumerate(HEADERS, start=1):
    ws.cell(row=2, column=column, value=header)

for row_number, values in enumerate(make_rows(), start=3):
    for column, value in enumerate(values, start=1):
        ws.cell(row=row_number, column=column, value=value)

ws.freeze_panes = "A3"
ws.auto_filter.ref = f"A2:U{ws.max_row}"
ws.row_dimensions[1].height = 26
ws.row_dimensions[2].height = 24
ws["A1"].font = ws["A2"].font.copy(bold=True, color="FFFFFF", size=14)
for cell in ws[2]:
    cell.font = ws["A1"].font.copy(size=10)
    cell.fill = ws["A1"].fill.copy()
    cell.alignment = ws["A1"].alignment.copy(horizontal="center")

for column in range(1, len(HEADERS) + 1):
    ws.column_dimensions[ws.cell(row=2, column=column).column_letter].width = 20
ws.column_dimensions["B"].width = 25
ws.column_dimensions["C"].width = 22
ws.column_dimensions["L"].width = 28
ws.column_dimensions["S"].width = 25

add_dropdown(ws, ",".join(TYPES), "D3:D122")
add_dropdown(ws, ",".join(FORM_FACTORS), "E3:E122")
add_dropdown(ws, ",".join(OEM_NAMES), "F3:F122")
add_dropdown(ws, ",".join(ENDURANCE), "I3:I122")
add_dropdown(ws, ",".join(PAYLOADS), "J3:J122")
add_dropdown(ws, ",".join(GUIDANCE), "M3:M122")
add_dropdown(ws, ",".join(ANTI_EW), "N3:N122")
add_dropdown(ws, ",".join(LINK_FREQ), "O3:O122")
add_dropdown(ws, ",".join(PROC_FUND), "P3:P122")
add_dropdown(ws, ",".join(SERVICE_STATUS), "Q3:Q122")
add_dropdown(ws, ",".join(UNITS), "T3:T122")
add_dropdown(ws, ",".join(FORMATIONS), "U3:U122")

ws.conditional_formatting.add(
    "R3:R122",
    ColorScaleRule(start_type="min", start_color="E7F3EC", mid_type="percentile", mid_value=50, mid_color="F5D48C", end_type="max", end_color="E98B74"),
)

OUTPUT.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUTPUT)
print(f"Wrote dummy workbook: {OUTPUT} ({len(make_rows())} records)")
