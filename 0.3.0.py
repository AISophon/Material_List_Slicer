#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Material List Slicer
Version 0.3.0
Author: Decent_Kook

This is a Python program for splitting a list of projected materials,
Users can customize weighted number, correlation coefficient and special item list.

Update Log
2025.01.17
-[Add]增加config配置文件
-[Add]没有配置文件自动新建，采用默认值
-[Opt]将提示与报错由print转为弹窗
[Add]
[Fix]
[Opt]
"""

from tkinter import filedialog,messagebox
from configparser import ConfigParser
import pandas as pd
import math
import os


def value_list_generating():  # configfile中读出的全为字符串类型，需要替换字符串中的特殊字符如“[”、“'”、“]”，然后重新生成列表类型。
    value = special_items.replace("[", "")  # 先替换所有的“[”为空，赋值给value
    value = value.replace("'", "")  # 再替换所有的“'”为空，再赋值给value
    value = value.replace("]", "")  # 再替换所有的“]”为空，再赋值给value
    value = value.replace(
        " ", ""
    )  # 再替换所有的空格为空，再赋值给value，至此字符串内特殊字符已全部删除
    value = value.split(",")  # 用“,”分割字符串，返回列表类数据。

    return value


# 配置文件路径
config_file = "config.ini"

# 创建 ConfigParser 对象
config = ConfigParser()

# 判断配置文件是否存在
if not os.path.exists(config_file):
    messagebox.showinfo("提示", "config.ini 不存在，正在创建默认配置文件...")

    # 添加 DEFAULT 部分并设置默认值
    config["DEFAULT"] = {
        "max_weighted_value": "8640",  # 单任务加权数最大值
        "normal_item_coefficient": "1",  # 普通物品系数
        "special_item_coefficient": "2",  # 特殊物品系数
        "item_type_coefficient": "500",  # 物品种类系数
        "special_items": "[]",  # 特殊物品
    }

    # 写入到配置文件
    with open(config_file, "w", encoding="utf-8") as configfile:
        config.write(configfile)

    messagebox.showinfo("提示", "config.ini 已创建，使用默认值。")
else:
    # 读取已存在的配置文件，并指定编码
    config.read(config_file, encoding="utf-8")

# 从配置文件中读取值
max_weighted_value = config["DEFAULT"].getint("max_weighted_value")
normal_item_coefficient = config["DEFAULT"].getint("normal_item_coefficient")
special_item_coefficient = config["DEFAULT"].getint("special_item_coefficient")
item_type_coefficient = config["DEFAULT"].getint("item_type_coefficient")
special_items = config["DEFAULT"].get("special_items", "[]")

special_items = value_list_generating()


# 读取csv文件
def read_file():
    file_path = filedialog.askopenfilename(
        title="请选择材料列表文件", filetypes=[("Excel", ".csv"), ("All Files", "*")]
    )
    return pd.read_csv(file_path, usecols=["Item", "Total"])


# 处理物品数量
def adjust_item_quantity(df):
    for row_index in range(len(df)):
        quantity = df.iloc[row_index, 1]
        if 1 < quantity <= 64:
            df.iat[row_index, 1] = math.ceil(
                ((128 - (quantity - 64) ** 2 / 32) // 16 + 1) * 16
            )
        elif 64 < quantity <= 576:
            df.iat[row_index, 1] = math.ceil(640 - (quantity - 576) ** 2 / 520)
        elif 576 < quantity:
            df.iat[row_index, 1] = (quantity // 576 + 1) * 576
    return df


# 获取物品的种类和数量，并强制转换为整数
def get_item_details(df, row_index):
    item_type = df.iloc[row_index, 0]
    quantity = int(df.iloc[row_index, 1])
    return item_type, quantity


# 配对任务
def pair_items_to_tasks(df):
    task_groups = []
    current_task = []
    total_weight = 0
    row_index = 0
    total_rows = len(df)

    # 获取物品种类和数量
    item_type, quantity = get_item_details(df, row_index)

    # 配对任务
    while row_index < total_rows:
        # 判断物品系数
        if item_type in special_items:
            item_coefficient = special_item_coefficient
        else:
            item_coefficient = normal_item_coefficient

        # 情况 1: 当前物品数量加权数小于目标加权数
        if max_weighted_value > quantity * item_coefficient + total_weight:
            current_task.append([item_type, quantity])
            total_weight += quantity * item_coefficient + item_type_coefficient

            # 判断是否满足加权数
            if total_weight >= max_weighted_value:
                task_groups.append(current_task)
                current_task = []
                total_weight = 0

            row_index += 1
            quantity_used = 0

            if row_index >= total_rows:
                task_groups.append(current_task)
                break

            item_type, quantity = get_item_details(df, row_index)

        # 情况 2: 当前物品数量加权数等于目标加权数
        elif max_weighted_value == quantity * item_coefficient + total_weight:
            current_task.append([item_type, quantity])
            task_groups.append(current_task)
            current_task = []
            row_index += 1
            total_weight = 0
            quantity_used = 0

            if row_index >= total_rows:
                break

            item_type, quantity = get_item_details(df, row_index)

        # 情况 3: 当前物品数量加权数大于目标加权数
        elif max_weighted_value < quantity * item_coefficient + total_weight:
            while max_weighted_value < quantity + total_weight:
                quantity -= 1
            if quantity <= 0:
                message = "quantity_need error: ", item_type, quantity
                messagebox.showerror("错误", message)
            current_task.append([item_type, quantity])
            task_groups.append(current_task)
            current_task = []
            total_weight = 0
            quantity_used += quantity
            quantity = int(df.iloc[row_index, 1]) - quantity_used

    return task_groups


# 展平数据并添加组别信息
def flatten_task_groups(task_groups):
    flattened_data = []
    for group_index, task_group in enumerate(task_groups, 1):
        for item in task_group:
            flattened_data.append([group_index] + item)
    return flattened_data


# 主程序逻辑
def main():
    df = read_file()
    df = adjust_item_quantity(df)
    task_groups = pair_items_to_tasks(df)
    flattened_data = flatten_task_groups(task_groups)

    # 创建 DataFrame，并设置列名
    result_df = pd.DataFrame(flattened_data, columns=["任务", "种类", "数量"])

    # 输出到Excel文件
    result_df.to_excel("output.xlsx", index=False)

    messagebox.showinfo("提示", "Excel 文件已成功保存为 output.xlsx")


# 执行主程序
if __name__ == "__main__":
    main()
