KUIN-G QR Generator — Input file format

The KUIN-G generator expects an Excel workbook with this minimal layout (the script validates this exact layout):

- Cell `B2` must contain the header: Drone ID
- Drone ID values must be placed in column B starting at row 3 (B3, B4, ...)
- Blank or empty values are ignored; each Drone ID produces one KUIN-G value

Example layout (visual):

     |   A           |  B
 1  |               |
   2  |               | Drone ID
   3  | 1             | CO-CO-DIV-BRIG-UNIT

Quick generator usage example:

python src/generate_UIDS.py \
  --input input/sample.xlsx \
  --master-output output/master_register.xlsx \
  --distributable-output output/distributable_register.xlsx \
  --qr-dir qr_codes

Notes:
- The code that enforces the layout is in `src/generate_UIDS.py`.
- Each KUIN-G QR is saved as `qr_codes/<KUIN-G>.png`.
- The tracking register is saved as `output/qr_register.csv`.
- If you need a sample file, run `tools/create_sample_input.py` to create `input/sample.xlsx`.
