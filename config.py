#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工作表导出为CSV工具
用于将指定Excel文件的指定工作表导出为CSV格式
使用方法: py config.py hero[hero]
"""

import os
import sys
from pathlib import Path
import pandas as pd
import openpyxl
import xlrd

# ==================== 配置区域 ====================
# 目标文件夹路径
TARGET_FOLDER = r"E:\qyn_game\parseFiles\global\config\test"

# 输出文件夹名称（在目标文件夹下创建）
OUTPUT_FOLDER = "xls"

# 支持的文件扩展名
SUPPORTED_EXTENSIONS = ['.xlsx', '.xls']
# ================================================

class ExcelToCSVConverter:
    def __init__(self, target_folder, output_folder):
        self.target_folder = Path(target_folder)
        # 输出文件夹在当前工作目录下，而不是目标文件夹下
        self.output_folder = Path.cwd() / output_folder

        # 确保输出文件夹存在
        self.output_folder.mkdir(exist_ok=True)
        
    def parse_command(self, command):
        """解析命令行参数，提取文件名和工作表名"""
        if '[' not in command or ']' not in command:
            raise ValueError("命令格式错误，应为: filename[sheetname]")

        # 分离文件名和工作表名
        file_part, sheet_part = command.split('[', 1)
        sheet_name = sheet_part.rstrip(']')

        if not file_part or not sheet_name:
            raise ValueError("文件名或工作表名不能为空")

        # 自动添加.xls扩展名
        filename = f"{file_part.strip()}.xls"

        return filename, sheet_name.strip()
    
    def find_excel_file(self, filename):
        """在目标文件夹中查找Excel文件"""
        file_path = self.target_folder / filename
        
        if file_path.exists():
            return file_path
        
        # 如果没有找到，尝试不同的扩展名
        name_without_ext = file_path.stem
        for ext in SUPPORTED_EXTENSIONS:
            test_path = self.target_folder / f"{name_without_ext}{ext}"
            if test_path.exists():
                return test_path
        
        return None
    
    def read_excel_sheet(self, file_path, sheet_name):
        """读取Excel文件的指定工作表"""
        file_extension = file_path.suffix.lower()
        
        try:
            if file_extension == '.xlsx':
                return self.read_xlsx_sheet(file_path, sheet_name)
            elif file_extension == '.xls':
                return self.read_xls_sheet(file_path, sheet_name)
            else:
                raise ValueError(f"不支持的文件格式: {file_extension}")
        except Exception as e:
            raise Exception(f"读取Excel文件失败: {str(e)}")
    
    def read_xlsx_sheet(self, file_path, sheet_name):
        """读取.xlsx文件的指定工作表"""
        try:
            # 首先检查工作表是否存在
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            if sheet_name not in workbook.sheetnames:
                available_sheets = ', '.join(workbook.sheetnames)
                raise ValueError(f"工作表 '{sheet_name}' 不存在。可用工作表: {available_sheets}")
            workbook.close()
            
            # 使用pandas读取指定工作表
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            return df
        except Exception as e:
            raise Exception(f"读取.xlsx文件失败: {str(e)}")
    
    def read_xls_sheet(self, file_path, sheet_name):
        """读取.xls文件的指定工作表"""
        try:
            # 首先检查工作表是否存在
            workbook = xlrd.open_workbook(file_path)
            sheet_names = workbook.sheet_names()
            if sheet_name not in sheet_names:
                available_sheets = ', '.join(sheet_names)
                raise ValueError(f"工作表 '{sheet_name}' 不存在。可用工作表: {available_sheets}")
            
            # 使用pandas读取指定工作表
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
            return df
        except Exception as e:
            raise Exception(f"读取.xls文件失败: {str(e)}")
    
    def save_to_csv(self, dataframe, output_filename):
        """将DataFrame保存为CSV文件"""
        output_path = self.output_folder / output_filename
        
        try:
            # 保存为CSV，使用UTF-8编码
            dataframe.to_csv(output_path, index=False, encoding='utf-8-sig')
            return output_path
        except Exception as e:
            raise Exception(f"保存CSV文件失败: {str(e)}")
    
    def convert(self, command):
        """执行转换操作"""
        try:
            # 解析命令
            filename, sheet_name = self.parse_command(command)
            print(f"解析命令: 文件名='{filename}', 工作表='{sheet_name}'")
            
            # 查找Excel文件
            file_path = self.find_excel_file(filename)
            if not file_path:
                raise FileNotFoundError(f"在目录 '{self.target_folder}' 中未找到文件 '{filename}'")
            
            print(f"找到文件: {file_path}")
            
            # 读取指定工作表
            print(f"正在读取工作表 '{sheet_name}'...")
            df = self.read_excel_sheet(file_path, sheet_name)
            
            print(f"成功读取数据: {len(df)} 行, {len(df.columns)} 列")
            
            # 生成输出文件名（去掉.xls扩展名）
            base_filename = Path(filename).stem  # 去掉扩展名
            output_filename = f"{base_filename}[{sheet_name}].csv"
            
            # 保存为CSV
            print(f"正在保存为CSV文件: {output_filename}")
            output_path = self.save_to_csv(df, output_filename)
            
            print(f"✅ 转换完成!")
            print(f"输出文件: {output_path}")
            print(f"数据行数: {len(df)}")
            print(f"数据列数: {len(df.columns)}")
            
            # 显示前几行数据预览
            if len(df) > 0:
                print("\n数据预览:")
                print(df.head().to_string())
            
        except Exception as e:
            print(f"❌ 转换失败: {str(e)}")
            return False
        
        return True

def main():
    """主函数"""
    print("Excel工作表导出为CSV工具")
    print("="*50)
    print(f"目标文件夹: {TARGET_FOLDER}")
    print(f"输出文件夹: {Path.cwd()}/{OUTPUT_FOLDER}")
    print("="*50)
    
    # 检查命令行参数
    if len(sys.argv) != 2:
        print("使用方法: py config.py filename[sheetname]")
        print("示例: py config.py hero[hero]")
        return
    
    command = sys.argv[1]
    print(f"执行命令: {command}")
    print()
    
    # 创建转换器并执行转换
    converter = ExcelToCSVConverter(TARGET_FOLDER, OUTPUT_FOLDER)
    success = converter.convert(command)
    
    if success:
        print("\n🎉 任务完成!")
    else:
        print("\n💥 任务失败!")

if __name__ == "__main__":
    main()
