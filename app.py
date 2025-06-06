import streamlit as st
import pandas as pd
import tempfile
import zipfile
import os
from tableaudocumentapi import Workbook
from xml.etree import ElementTree as ET
import re

st.set_page_config(page_title="Tableau Metadata Extractor", layout="wide")
st.title("📊 Tableau TWBX Metadata Extractor")

uploaded_file = st.file_uploader("Upload a Tableau TWBX or TWB file", type=["twbx", "twb"])

def resolve_formula(formula, calc_map, visited):
    if not formula or formula == "-":
        return formula
    matches = re.findall(r"\[([^\]]+)\]", formula)
    for m in matches:
        key = m.strip()
        if key in visited:
            continue
        if key in calc_map:
            visited.add(key)
            resolved = resolve_formula(calc_map[key]['calculation'], calc_map, visited)
            formula = formula.replace(f"[{key}]", f"({resolved})")
        elif key in calc_map:
            formula = formula.replace(f"[{key}]", f"[{calc_map[key]['caption']}]")
    return formula

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_file_path = os.path.join(tmpdir, uploaded_file.name)
        with open(temp_file_path, "wb") as f:
            f.write(uploaded_file.read())

        if uploaded_file.name.endswith(".twbx"):
            with zipfile.ZipFile(temp_file_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
                twb_files = [f for f in zip_ref.namelist() if f.endswith(".twb")]
                if not twb_files:
                    st.error("No .twb file found inside TWBX.")
                    st.stop()
                twb_path = os.path.join(tmpdir, twb_files[0])
        else:
            twb_path = temp_file_path

        wb = Workbook(twb_path)

        # Build calculation map
        calc_map = {}
        for ds in wb.datasources:
            for field in ds.fields.values():
                calc_map[field.name] = {
                    'caption': field.caption or field.name,
                    'calculation': field.calculation if field.calculation is not None else "-"
                }

        # Extract all fields from datasources
        all_fields_data = []
        for ds in wb.datasources:
            for field in ds.fields.values():
                raw_formula = field.calculation if field.calculation is not None else "-"
                resolved_formula = resolve_formula(raw_formula, calc_map, visited=set()) if raw_formula != "-" else "-"

                all_fields_data.append({
                    "Datasource": ds.caption,
                    "Field Name": field.caption or field.name,
                    "Data Type": field.datatype,
                    "Role": field.role,
                    "Is Calculated": field.calculation is not None,
                    "Formula": resolved_formula
                })
        all_fields_df = pd.DataFrame(all_fields_data)

        xml_root = ET.parse(twb_path).getroot()
        worksheet_data = []
        for worksheet_elem in xml_root.findall(".//worksheet"):
            ws_name = worksheet_elem.get("name")
            filters = [f.get("column") for f in worksheet_elem.findall(".//filter") if f.get("column")]

            rows_fields = []
            cols_fields = []
            for pane in worksheet_elem.findall(".//pane"):
                for shelf in pane.findall("shelf"):
                    shelf_type = shelf.get("type")
                    fields = [f.get("field") for f in shelf.findall(".//encoding") if f.get("field")]
                    if shelf_type == "rows":
                        rows_fields.extend(fields)
                    elif shelf_type == "columns":
                        cols_fields.extend(fields)

            worksheet_data.append({
                "Worksheet Name": ws_name,
                "Filters Used": ", ".join(filters) if filters else "-",
                "Fields on Rows Shelf": ", ".join(rows_fields) if rows_fields else "-",
                "Fields on Columns Shelf": ", ".join(cols_fields) if cols_fields else "-"
            })
        worksheets = pd.DataFrame(worksheet_data)

        dashboard_elems = xml_root.findall('.//dashboard')
        dashboard_names = [db.get('name') for db in dashboard_elems if db.get('name')]
        dashboards = pd.DataFrame(dashboard_names, columns=["Dashboard Name"])

        ds_data = []
        for ds in wb.datasources:
            for conn in ds.connections:
                ds_data.append({
                    "Datasource Name": ds.caption,
                    "Connection Type": conn.dbclass or "Unknown",
                    "Server": getattr(conn, 'server', '-'),
                    "Database": getattr(conn, 'dbname', '-'),
                    "Table": getattr(conn, 'relation_name', '-'),
                    "Username": getattr(conn, 'username', '-'),
                    "Uses Extract": getattr(conn, 'is_extract', False)
                })
        datasources = pd.DataFrame(ds_data)

        st.header("📄 Worksheets")
        st.dataframe(worksheets)

        st.header("📊 Dashboards")
        st.dataframe(dashboards)

        st.header("🧩 Data Sources")
        st.dataframe(datasources)

        st.header("🧾 Fields in Data Sources")
        st.dataframe(all_fields_df)

        st.success("Metadata extraction complete ✅")
