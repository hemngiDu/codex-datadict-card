"""Create a small synthetic Kingdee export; contains no customer data."""
from pathlib import Path
import sys
import openpyxl


def create(path):
    wb = openpyxl.Workbook()
    catalog = wb.active
    catalog.title = "表目录"
    specs = [
        ("采购订单", "t_pm_purorder", "", [("fid", "主键", True), ("fsupplierid", "供应商", False)]),
        ("采购订单", "t_pm_purorderentry", "t_pm_purorder", [("fentryid", "分录主键", True), ("fid", "采购订单", False), ("fqty", "采购数量", False)]),
        ("供应商", "t_bd_supplier", "", [("fid", "主键", True), ("fname", "供应商名称", False)])]
    for i, (bo, name, parent, cols) in enumerate(specs, 7):
        for col, value in ((4, i-6), (5, bo), (6, bo), (7, name), (8, parent)):
            catalog.cell(i, col, value)
        ws = wb.create_sheet(str(i))
        ws.append([None, "表名：", None, None, bo])
        ws.append([None, "数据表名：", None, None, name])
        ws.append([None, "编号", "列名", "列标题", "类型", None, None, "主键"])
        for n, (field, desc, pk) in enumerate(cols, 1):
            ws.append([None, n, field, desc, "varchar" if field == "fname" else "bigint", None, None, "是" if pk else "否"])
    wb.save(path)
    wb.close()


if __name__ == "__main__":
    target = Path(sys.argv[1])
    target.parent.mkdir(parents=True, exist_ok=True)
    create(target)
