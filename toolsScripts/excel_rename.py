"""
读取Excel文件，去除第一列中每个单元格"_"及其后面的内容，并保存为新文件（<filename>-修改后.xlsx）。
"""

import os
import openpyxl

if __name__ == '__main__':
    filename = input("请输入要处理的Excel文件名：").strip()

    if os.path.splitext(filename)[1] == "":
        filename += ".xlsx"
    elif os.path.splitext(filename)[1].lower() != ".xlsx":
        print("目前仅支持.xlsx格式的Excel文件。")
        input("按回车键退出。")
        exit(1)

    if os.path.exists(filename):
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active

        for row in sheet.iter_rows():
            if row[0].value is not None and "_" in row[0].value:
                print(f"{row[0].value} -> {row[0].value.split('_')[0]}")
                row[0].value = row[0].value.split("_")[0]

        wb.save(f"{os.path.basename(filename)}-修改后.xlsx")
        wb.close()

    else:
        print(f"文件 {filename} 不存在，请检查后重试。")

    input("处理完成，按回车键退出。")
