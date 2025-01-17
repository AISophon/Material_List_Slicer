#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Material List Slicer
Version 0.2.4
Author: Decent_Kook

This is a Python program for splitting a list of projected materials,
Users can customize weighted number, correlation coefficient and special item list.

Update Log
2025.01.16
-[Opt]对变量名进行规范化和易读化处理
"""

from tkinter import filedialog
import pandas as pd
import math

# 设定划分标准（用户自行决定）
max_weighted_value = 1728 * 5  # 单任务加权数最大值
normal_item_coefficient = 1  # 普通物品系数
special_item_coefficient = 2  # 特殊物品系数
item_type_coefficient = 500  # 物品种类系数
special_items = []  # 特殊物品


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
                print("quantity_need error: ", item_type, quantity)
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

    print("Excel 文件已成功保存为 output.xlsx")


# 执行主程序
if __name__ == "__main__":
    main()
