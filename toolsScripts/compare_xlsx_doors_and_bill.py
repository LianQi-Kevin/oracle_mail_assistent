"""交叉对比门窗表和工程量清单"""
import os

import openpyxl
import re
from dataclasses import dataclass, field

from main import clean_str

DOOR_SHEET = r"./门窗表.xlsx"
BILL_SHEET = r"./工程量清单.xlsx"


@dataclass
class doorData:
    name: str
    facing: str = field(default="N/A")
    window: bool = field(default=False)
    num: int = field(default=0)


def load_door_data() -> dict[str, doorData]:
    """加载门窗表数据"""
    door_data_: dict[str, doorData] = dict()

    print(f"正在处理门窗表: {DOOR_SHEET}")

    wb = openpyxl.load_workbook(DOOR_SHEET, data_only=True, read_only=True)

    ws_list = [("地下室", 7, 9), ("人防出入口", 4, 6)]

    for sheet_name, col_index, facing_index in ws_list:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[1] is not None and row[col_index] is not None and int(row[col_index]) > 0:
                cleaned_key = clean_str(row[1])

                # 构造doorData对象
                door_data_[cleaned_key] = doorData(
                    name=cleaned_key, num=int(row[col_index]),
                    facing=row[facing_index] if "/" not in row[facing_index] else row[facing_index].split("/")[0].strip(), window=("观察窗" in row[facing_index]))

    # 检查door_data_中的是否存在重复
    duplicates = set([x for x in door_data_.keys() if list(door_data_.keys()).count(x) > 1])
    if duplicates:
        print(f"门窗表: {os.path.basename(DOOR_SHEET)}中存在重复项: {duplicates}")

    return door_data_


def load_bill_data(bill_sheet: str) -> dict[str, doorData]:
    """加载工程量清单数据"""
    print(f"正在处理工程量清单: {bill_sheet}")

    bill_data_: dict[str, doorData] = dict()

    wb = openpyxl.load_workbook(bill_sheet, data_only=True, read_only=True)
    ws = wb["地下建筑工程-门窗"]

    key_list = []

    for row in ws.iter_rows(min_row=6, values_only=True):
        if row[3] is not None and row[6] is not None and int(row[6]) > 0:
            # 使用正则表达式提取cleaned_str中中文前的所有非中文字符
            cleaned_key = clean_str(row[3]).split(":")[1].strip()
            match = re.match(r'^([^\u4e00-\u9fa5]*)', cleaned_key)
            cleaned_key = match.group(1).split(" ")[0] if match else ""
            if cleaned_key.endswith("("):  # 清理紧贴括号的情况
                cleaned_key = cleaned_key[:-1]

            key_list.append(cleaned_key)

            # 构造doorData对象
            bill_data_[cleaned_key] = doorData(
                name=cleaned_key, num=int(row[6]),
                # facing=row[5] if row[5] and "/" not in row[5] else (row[5].split("/")[0].strip() if row[5] else None),
                window=("观察窗" in row[3]))

    # 检查key_list中的是否存在重复
    duplicates = set([x for x in key_list if key_list.count(x) > 1])
    if duplicates:
        print(f"工程量清单: {os.path.basename(bill_sheet)}中存在重复项: {duplicates}")

    return bill_data_


@dataclass
class cleanDictItem:
    # wb_path: str
    idx: int
    door_name: str
    sheet_name: str


if __name__ == '__main__':
    door_data = load_door_data()

    bill_data: dict[str, doorData] = dict()

    for _bill_sheet in BILL_SHEET:
        bill_data.update(load_bill_data(_bill_sheet))

    # 交叉对比两个dict
    all_keys = set(door_data.keys()).union(set(bill_data.keys()))
    for key in all_keys:
        door_entry = door_data.get(key)
        bill_entry = bill_data.get(key)
        door_qty = door_entry.num if door_entry else 0
        bill_qty = bill_entry.num if bill_entry else 0
        if door_qty != bill_qty:
            door_facing = door_entry.facing if door_entry else "N/A"
            door_window = "有" if door_entry and door_entry.window else "无"
            bill_facing = bill_entry.facing if bill_entry else "N/A"
            bill_window = "有" if bill_entry and bill_entry.window else "无"
            print(f"项目: {key} | 门窗表数量: {door_qty} (饰面: {door_facing}, 观察窗: {door_window}) | "
                  f"工程量清单数量: {bill_qty} (饰面: {bill_facing}, 观察窗: {bill_window}) | "
                  f"差异: {door_qty - bill_qty}")
        else:
            print(f"项目: {key} 数量一致: {door_qty}, 饰面: {door_entry.facing if door_entry else 'N/A'}, 观察窗: {'有' if door_entry and door_entry.window else '无'}")
            # pass
    print("数量交叉对比完成")

    total_itemized_dict: dict[str, cleanDictItem] = dict()

    for _bill_sheet in BILL_SHEET:
        itemized_dict: dict[str, cleanDictItem] = dict()

        wb = openpyxl.load_workbook(_bill_sheet, read_only=True, data_only=True)
        ws_list = [clean_str(sheet.title).replace(" ", "") for sheet in wb.worksheets]
        for idx, sheet in [(i, sheet) for i, sheet in enumerate(wb.worksheets) if not bool(re.search(r'[\u4e00-\u9fa5]', sheet.title))]:
            cleaned_door_name = clean_str(sheet.cell(row=8, column=2).value).replace(" ", "")
            itemized_dict[cleaned_door_name] = cleanDictItem(idx=idx, door_name=cleaned_door_name, sheet_name=sheet.title)

        print(f"工程量清单: {os.path.basename(_bill_sheet)}中分项有: {list(itemized_dict.keys())}")

        total_itemized_dict.update(itemized_dict)

    # 对比total_itemized_dict.keys()和door_data.keys()
    all_keys = set(total_itemized_dict.keys()).union(set(door_data.keys()))
    for key in all_keys:
        itemized_entry = total_itemized_dict.get(key)
        door_entry = door_data.get(key)
        if door_entry and not itemized_entry:
            print(f"门窗表项目: {door_entry.name} 在工程量清单分项中未找到对应项")
