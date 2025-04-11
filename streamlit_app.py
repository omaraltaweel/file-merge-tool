import streamlit as st
import pandas as pd
from datetime import datetime
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

EXPECTED_HEADERS = [
    "Material_ID", "Description", "Buyer_Group", "Technical Text", "Standard_Material_Set", "COPIC_Number",
    "NIIN", "Manufacturing_Part_No", "Part_No", "Maturity", "ITAR", "TypeDescription", "MaterialType",
    "ExternalMaterialStatus", "StandardMaterialCategory", "TechnicalResponsibleUser", "MatSpec", "GFX",
    "NatoStockID", "ProcPack", "Weight", "Height", "Width", "Depth", "StandardMaterialClass", "CageCode",
    "ContPartNo", "ContPartName", "ContNo", "ProcClass", "EAR600", "EAR", "NonStdRtl", "Unit", "Certificate",
    "StockShelfLife", "HazardousMaterial", "Equivalentmaterial", "machined", "machiningStrategy", "BuildPhaseID",
    "LockoutLoadout", "hotwork", "UnitBlockBreak", "Flushed", "PipeInstallationTestMedium", "PipeInstallationTestPressure",
    "PipeShopTestMedium", "PipeShopTestPressure", "PipeFlushingMedium", "PipeFlushingAcceptanceCriteria",
    "PipeAdditionalTestPressCrit", "PipeAdditionalTestMedium", "AuthoringApplication", "Identifier", "Interface",
    "Inspection_Codes", "CommodityCode", "Min_Order_Qty", "Max_Order_Qty", "Supplier_ID"
]

UNWANTED_COLUMNS = [
    "ITAR", "TypeDescription", "StandardMaterialClass", "Certificate",
    "StockShelfLife", "AuthoringApplication", "Identifier", "Interface"
]

st.title("📊 FileMergeTool")
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    validation_errors = {}
    warnings = []

    for file in uploaded_files:
        try:
            xl = pd.ExcelFile(file)
            if "Standard Materials" not in xl.sheet_names:
                validation_errors[file.name] = ["Missing 'Standard Materials' sheet"]
                continue

            df = xl.parse("Standard Materials", dtype=str).fillna("")
            df.columns = [str(c).strip() for c in df.columns]

            errors = []
            for i, expected in enumerate(EXPECTED_HEADERS):
                if i >= len(df.columns):
                    errors.append(f"Missing column {i+1}: '{expected}'")
                elif df.columns[i] != expected:
                    errors.append(f"Column {i+1}: Found '{df.columns[i]}', expected '{expected}'")

            if errors:
                validation_errors[file.name] = errors
                continue

            df.insert(0, "S.I.", "")
            df["SourceFile"] = file.name
            all_data.append(df)

        except Exception as e:
            validation_errors[file.name] = [str(e)]

    if validation_errors:
        st.error("Validation Errors Found:")
        for f, msgs in validation_errors.items():
            st.write(f"**{f}**")
            for msg in msgs:
                st.write(f" - {msg}")
        st.stop()

    st.success(f"{len(all_data)} files passed validation. Merging...")

    combined_df = pd.concat(all_data, ignore_index=True)

    # Handle formatting
    buffer = io.BytesIO()
    combined_df.to_excel(buffer, index=False, sheet_name="Merged Data")
    buffer.seek(0)
    wb = load_workbook(buffer)
    ws = wb["Merged Data"]

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    material_ids = [cell.value for cell in ws["B"][1:]]
    duplicates = {x for x in material_ids if material_ids.count(x) > 1}
    for cell in ws["B"][1:]:
        if cell.value in duplicates:
            cell.fill = red_fill

    # Handle unwanted columns
    header_row = [cell.value for cell in ws[1]]
    deleted_cols = []
    non_empty_highlight_cols = []

    for col_idx, header in enumerate(header_row, start=1):
        if header in UNWANTED_COLUMNS:
            col_data = list(ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, values_only=True))[0]
            if all(v in ("", None) for v in col_data):
                deleted_cols.append(col_idx)
            else:
                non_empty_highlight_cols.append(header)

    for col_idx in sorted(deleted_cols, reverse=True):
        ws.delete_cols(col_idx)

    for col_idx, header in enumerate([cell.value for cell in ws[1]], start=1):
        if header in UNWANTED_COLUMNS:
            col_data = list(ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, values_only=True))[0]
            if any(v not in ("", None) for v in col_data):
                ws.cell(row=1, column=col_idx).fill = red_fill

    max_row, max_col = ws.max_row, ws.max_column
    ws.add_table(Table(displayName="MergedDataTable", ref=f"A1:{get_column_letter(max_col)}{max_row}",
                       tableStyleInfo=TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)))
    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 2, 50)

    # Save to buffer for download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    if non_empty_highlight_cols:
        warnings.append("⚠️ Some unwanted columns could not be deleted because they contain data:\n- " + "\n- ".join(non_empty_highlight_cols))
    if duplicates:
        warnings.append("⚠️ Duplicate item IDs were found and highlighted in red.")

    if warnings:
        st.warning("\n\n".join(warnings))

    st.success("Merge complete!")
    st.download_button("📥 Download Merged Excel", output, file_name="MergedOutput.xlsx")

