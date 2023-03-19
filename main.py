#!/usr/bin/env python
# -*- coding:UTF-8 -*-
#
# @DESCRIPTION: 为 Excel 表加上二维码一列


import os
import qrcode
import xlrd
import xlwt
from PIL import Image


# QRCode 存储路径
QRCode_Path = "./QRcode"


def generate_qrcode_img(data: str):
    """
    生成二维码图片
    """
    # 1. 生成 .png 文件
    qrcode_img = qrcode.make(data)
    qrcode_img.save(f"{QRCode_Path}/{data}.png")
    # 2. 转为 .bmp 文件
    with Image.open(f"{QRCode_Path}/{data}.png") as qrcode_img:
        img_rgb = qrcode_img.convert('RGB')
        img_24 = Image.new('RGB', img_rgb.size, (255, 255, 255))
        img_24.paste(img_rgb, mask=img_rgb.split()[3] if len(img_rgb.split()) == 4 else None)
        # 将大小缩小到 150*150
        img_24 = img_24.resize((150, 150))
        # 将质量减少至 1/10
        img_24.save(f"{QRCode_Path}/{data}.bmp", quality=10)
    return f"{QRCode_Path}/{data}.bmp"


def main():
    """
    主函数
    """
    # 0. 如果目录下没有 QRCode 文件夹的话，创建
    if not os.path.exists(QRCode_Path):
        os.mkdir(QRCode_Path)
    # 1. 读取目录下的所有 Excel，Excel 表以 .xlsx 和 .xls 结尾
    file_names = os.listdir("./")
    excel_file_names = []
    for file_name in file_names:
        if (file_name.endswith(".xlsx") or file_name.endswith(".xls"))\
                and not file_name.startswith("带二维码的"):
            excel_file_names.append(file_name)
    print("有 {} 个 Excel 表".format(len(excel_file_names)))
    # 2. 遍历所有 Excel
    for excel_file_name in excel_file_names:
        # 2.1 读取 Excel 并选取活跃的 Sheet
        excel = xlrd.open_workbook(excel_file_name)
        sheet = excel.sheet_by_index(0)
        # 2.2 创建新的 Excel 用一写入
        new_excel = xlwt.Workbook()
        new_sheet = new_excel.add_sheet("带二维码的{}".format(sheet.name))
        # 2.3 复制原 Excel 的内容到新 Excel
        for row in range(0, sheet.nrows):
            for col in range(0, sheet.ncols):
                new_sheet.write(row, col, sheet.cell_value(row, col))
        # 2.5 设置列宽为 20、30 和 30
        new_sheet.col(0).width = 20 * 256
        new_sheet.col(1).width = 30 * 256
        new_sheet.col(sheet.ncols).width = 25*256
        # 2.4 在新 Excel 的最后一列加上二维码
        new_sheet.write(0, sheet.ncols, "二维码")
        for row in range(1, sheet.nrows):
            # 2.5 设置行高
            new_sheet.row(row).height_mismatch = True
            new_sheet.row(row).height = 80 * 20
            # 2.6.1 生成二维码图片
            qrcode_img_path = generate_qrcode_img(sheet.cell_value(row, 0))
            with open(qrcode_img_path, "rb") as qrcode_img_file:
                # 2.6.2 将二维码图片写入到新 Excel
                new_sheet.insert_bitmap(qrcode_img_path, row, sheet.ncols, scale_x=0.55, scale_y=0.11, x=1, y=1)
            # 2.7 在最后一列写 6 个回车
            new_sheet.write(row, sheet.ncols, "\r" * 4)
            # 2.8 打印进度
            print("已完成 {} / {} 行".format(row, sheet.nrows))
        # 2.9 保存新 Excel
        new_excel.save(f"带二维码的{excel_file_name}")


# 启动入口
if __name__ == '__main__':
    main()