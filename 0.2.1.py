#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Material List Slicer
Version 0.2.1
Author: Decent_Kook

This is a Python program for splitting a list of projected materials,
Users can customize weighted number, correlation coefficient and special item list.

Update Log
2025.01.16
-[Add]弹出窗口选择材料列表文件
-[Fix]优化取整逻辑
"""

from tkinter import filedialog
import pandas as pd
import math

# 设定划分标准（用户自行决定）
weighted_number = 1728 * 5  # 单任务加权数最大值
ordinary_coefficient = 1  # 普通物品系数
specific_coefficient = 2  # 特殊物品系数
kind_coefficient = 500  # 物品种类系数
specific_item = []  # 特殊物品

# 读取csv文件
path = filedialog.askopenfilename(
    title="请选择材料列表文件", filetypes=[("Excel", ".csv"), ("All Files", "*")]
)
df = pd.read_csv(path, usecols=["Item", "Total"])

# 处理列表
quantity_used = 0
row_index = 0
row_max = len(df)
output = []
task = []
sum = 0

# 取整
while row_index < row_max:
    # 获取物品种类和数量
    kind = df.iloc[row_index, 0]  # 物品种类
    quantity = df.iloc[row_index, 1]  # 物品数量
    if 1 < quantity <= 64:
        df.iat[row_index, 1] = math.ceil(
            ((128 - (quantity - 64) ** 2 / 32) // 16 + 1) * 16
        )
    elif 64 < quantity <= 576:
        df.iat[row_index, 1] = math.ceil(640 - (quantity - 576) ** 2 / 520)
    elif 576 < quantity:
        df.iat[row_index, 1] = (quantity // 576 + 1) * 576
    row_index += 1
row_index = 0

# 获取物品种类和数量
kind = df.iloc[row_index, 0]  # 物品种类
quantity = df.iloc[row_index, 1]  # 物品数量

# 强制转换 np.int64 为 int
quantity = int(quantity)

# 配对任务
while row_index < row_max:
    # 判断物品系数
    if kind in specific_item:
        coefficient = specific_coefficient
    else:
        coefficient = ordinary_coefficient

    if weighted_number > quantity * coefficient + sum:
        task.append([kind, quantity])

        # 计算加权数
        sum += quantity * coefficient + kind_coefficient
        if sum >= weighted_number:
            output.append(task)
            task = []
            sum = 0
        row_index += 1
        quantity_used = 0

        # 获取物品种类和数量
        if row_index >= row_max:
            output.append(task)
            break
        kind = df.iloc[row_index, 0]  # 物品种类
        quantity = df.iloc[row_index, 1]  # 物品数量

        # 强制转换 np.int64 为 int
        quantity = int(quantity)

    elif weighted_number == quantity * coefficient + sum:
        task.append([kind, quantity])
        output.append(task)
        task = []
        row_index += 1
        sum = 0
        quantity_used = 0

        # 获取物品种类和数量
        if row_index >= row_max:
            break
        kind = df.iloc[row_index, 0]  # 物品种类
        quantity = df.iloc[row_index, 1]  # 物品数量

        # 强制转换 np.int64 为 int
        quantity = int(quantity)
    elif weighted_number < quantity * coefficient + sum:
        while weighted_number < quantity + sum:
            quantity -= 1
        if quantity <= 0:
            print("quantity_need error: ", kind, quantity)
        task.append([kind, quantity])
        output.append(task)
        task = []
        sum = 0
        quantity_used = quantity_used + quantity
        quantity = int(df.iloc[row_index, 1]) - quantity_used

# 输出Excel
flattened_data = []
for sublist in output:
    for item in sublist:
        flattened_data.append(item)

# 用于存储展平后的数据
flattened_data = []

# 展平数据并添加组别信息
for idx, sublist in enumerate(output, 1):  # 使用enumerate为每个子列表添加组编号
    for item in sublist:
        # 将每个数据项和它所属的组号一起存储
        flattened_data.append([idx] + item)

# 创建 DataFrame，并设置列名
df = pd.DataFrame(flattened_data, columns=["任务", "种类", "数量"])

# 输出到Excel文件
df.to_excel("output.xlsx", index=False)

print("Excel 文件已成功保存为 output.xlsx")
