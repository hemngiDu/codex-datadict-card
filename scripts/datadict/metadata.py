from pathlib import Path
import openpyxl


def read_excel(path):
    """Read the Kingdee catalog and every column without truncation."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    tables = {}
    warnings = []
    def text(value):
        return str(value).strip() if value is not None else ""
    try:
        if "表目录" not in wb.sheetnames:
            raise ValueError("Missing 表目录 worksheet")
        for sheet in wb:
            if sheet.max_row == 1 and sheet.max_column == 1:
                sheet.reset_dimensions()
        bo = ""
        for row in wb["表目录"].iter_rows(min_row=7, max_col=11, values_only=True):
            bo = text(row[4]) or bo
            name = text(row[6]).lower()
            if name and row[5]:
                ref = dict(business_name=bo, description=text(row[5]), parent=text(row[7]).lower())
                if name not in tables:
                    tables[name] = dict(name=name, **ref, references=[], columns=[])
                if ref not in tables[name]["references"]:
                    tables[name]["references"].append(ref)
        for ws in wb:
            if ws.title == "表目录":
                continue
            current, in_columns = None, False
            for row in ws.iter_rows(max_col=11, values_only=True):
                marker = text(row[1])
                if marker == "表名：":
                    current, in_columns = None, False
                elif marker == "数据表名：":
                    current = tables.get(text(row[4]).lower())
                    in_columns = False
                elif marker == "编号" and text(row[2]) == "列名":
                    in_columns = True
                elif text(row[7]) == "返回主目录":
                    current, in_columns = None, False
                elif current is not None and in_columns and marker:
                    try:
                        float(marker)
                    except ValueError:
                        continue
                    name = text(row[2]).lower()
                    if not name:
                        continue
                    column = dict(name=name, description=text(row[3]),
                        type=text(row[4]).lower(), length=text(row[5]), scale=text(row[6]),
                        nullable_marker=text(row[8]), default=text(row[9]), primary_key=text(row[7]).lower() in
                        {"是", "√", "true", "1", "yes", "y", "pk"}, remarks=text(row[10]))
                    existing = next ( (c for c in current["columns"] if c["name"] == name), None)
                    if existing is None:
                        current["columns"].append(column)
                    elif existing != column:
                        conflict = dict(sheet=ws.title, table=current["name"], column=column)
                        if conflict not in warnings: warnings.append(conflict)
        if not tables:
            raise ValueError("No tables found in catalog")
        return dict(schema_version=1, source=Path(path).name, tables=list(tables.values()), metadata_conflicts=warnings)
    finally:
        wb.close()
