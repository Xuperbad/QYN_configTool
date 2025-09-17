#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工作表导出为CSV工具
用于将指定Excel文件的指定工作表导出为CSV格式
使用方法: py config.py hero[hero]
"""

import sys
from pathlib import Path
import pandas as pd
import openpyxl
import xlrd
import xlwt
import re
import multiprocessing
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

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
    
    def find_t_strings(self, text):
        """查找文本中所有的t_*字符串，先按逗号分割再匹配"""
        if pd.isna(text) or not isinstance(text, str):
            return []

        t_strings = []

        # 先按逗号分割文本
        parts = [part.strip() for part in text.split(',')]

        for part in parts:
            # 检查每个部分是否是完整的t_*字符串
            if re.match(r'^t_[a-zA-Z0-9_]+$', part):
                t_strings.append(part)

        return list(set(t_strings))  # 去重

    def search_chinese_text(self, t_string):
        """直接调用go.py的方法获取t_string对应的中文文本"""
        try:
            # 导入go.py模块
            import sys
            sys.path.append(str(Path.cwd()))
            from go import ExcelTextReplacer, TARGET_FOLDER

            # 创建ExcelTextReplacer实例
            replacer = ExcelTextReplacer({})  # 空的替换配置，因为我们只是用来搜索

            # 直接调用新的方法获取中文文本
            chinese_text = replacer.get_chinese_text_by_id(t_string, TARGET_FOLDER)

            return chinese_text

        except Exception as e:
            # 如果出现任何错误，返回None
            return None

    def search_chinese_text_batch(self, t_strings, max_workers=None):
        """并发批量搜索t_string对应的中文文本"""
        if max_workers is None:
            # 根据CPU核心数和任务数量动态调整线程数
            cpu_count = multiprocessing.cpu_count()
            # 使用CPU核心数的2-3倍，但不超过任务数量，最少4个线程
            max_workers = min(max(cpu_count * 3, 4), len(t_strings), 32)

        print(f"使用 {max_workers} 个线程并发搜索...")

        results = {}
        completed_count = 0
        total_count = len(t_strings)
        lock = threading.Lock()

        def search_single(t_string):
            nonlocal completed_count
            chinese_text = self.search_chinese_text(t_string)

            with lock:
                completed_count += 1
                if chinese_text:
                    print(f"  [{completed_count}/{total_count}] {t_string} -> {chinese_text}")
                else:
                    print(f"  [{completed_count}/{total_count}] {t_string} -> 未找到")

            return t_string, chinese_text

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有任务
            future_to_string = {executor.submit(search_single, t_string): t_string
                              for t_string in t_strings}

            # 收集结果
            for future in as_completed(future_to_string):
                try:
                    t_string, chinese_text = future.result()
                    results[t_string] = chinese_text
                except Exception as e:
                    t_string = future_to_string[future]
                    print(f"搜索 {t_string} 时出错: {str(e)}")
                    results[t_string] = None

        return results

    def preprocess_dataframe(self, df):
        """预处理DataFrame，将t_*字符串替换为t_*{中文}格式"""
        print("正在进行预处理，识别并查找t_*字符串...")

        # 收集所有需要查找的t_*字符串
        all_t_strings = set()

        for col in df.columns:
            for idx, cell_value in df[col].items():
                if pd.notna(cell_value) and isinstance(cell_value, str):
                    t_strings = self.find_t_strings(cell_value)
                    all_t_strings.update(t_strings)

        if not all_t_strings:
            print("未找到任何t_*字符串，跳过预处理")
            return df

        print(f"找到 {len(all_t_strings)} 个唯一的t_*字符串，正在并发查找对应中文...")

        # 并发批量查找中文文本
        chinese_results = self.search_chinese_text_batch(list(all_t_strings))

        # 构建替换映射
        t_string_map = {}
        found_count = 0
        for t_string, search_result in chinese_results.items():
            if search_result:
                # search_result 现在只包含中文内容
                t_string_map[t_string] = f"{t_string}{{{search_result}}}"
                found_count += 1
            else:
                t_string_map[t_string] = t_string  # 保持原样（搜索结果不唯一或未找到）

        # 替换DataFrame中的内容
        print("正在替换DataFrame中的内容...")
        df_processed = df.copy()

        for col in df_processed.columns:
            for idx, cell_value in df_processed[col].items():
                if pd.notna(cell_value) and isinstance(cell_value, str):
                    new_value = cell_value

                    # 按逗号分割，对每个部分单独处理
                    parts = [part.strip() for part in new_value.split(',')]
                    new_parts = []

                    for part in parts:
                        # 检查这个部分是否是完整的t_*字符串
                        if part in t_string_map:
                            new_parts.append(t_string_map[part])
                        else:
                            new_parts.append(part)

                    df_processed.loc[idx, col] = ', '.join(new_parts)

        print(f"预处理完成，共找到 {found_count}/{len(all_t_strings)} 个t_*字符串的中文对应")
        print(f"替换了 {found_count} 个t_*字符串为带中文的格式")
        return df_processed

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

            # 预处理DataFrame，查找并替换t_*字符串
            df_processed = self.preprocess_dataframe(df)

            # 生成输出文件名（去掉.xls扩展名）
            base_filename = Path(filename).stem  # 去掉扩展名
            output_filename = f"{base_filename}[{sheet_name}].csv"

            # 保存为CSV
            print(f"正在保存为CSV文件: {output_filename}")
            output_path = self.save_to_csv(df_processed, output_filename)

            print(f"✅ 转换完成!")
            print(f"输出文件: {output_path}")
            print(f"数据行数: {len(df_processed)}")
            print(f"数据列数: {len(df_processed.columns)}")

            # 显示前几行数据预览
            if len(df_processed) > 0:
                print("\n数据预览:")
                print(df_processed.head().to_string())

        except Exception as e:
            print(f"❌ 转换失败: {str(e)}")
            return False

        return True

    def process_csv_content(self, csv_content):
        """处理CSV内容，将t_*{中文}格式还原为t_*格式"""
        import re

        # 正则表达式匹配 t_*{中文} 格式
        pattern = r't_([a-zA-Z0-9_]+)\{[^}]*\}'

        # 替换为 t_* 格式
        processed_content = re.sub(pattern, r't_\1', csv_content)

        return processed_content

    def write_csv_to_excel(self, csv_file_path, excel_file_path, sheet_name):
        """将CSV文件内容写回到Excel文件的指定工作表"""
        try:
            # 读取CSV文件
            print(f"正在读取CSV文件: {csv_file_path}")
            df = pd.read_csv(csv_file_path, encoding='utf-8-sig')

            # 处理CSV内容，将t_*{中文}格式还原为t_*格式
            print("正在处理CSV内容，还原t_*字符串格式...")
            for col in df.columns:
                df[col] = df[col].astype(str).apply(lambda x: self.process_csv_content(x) if pd.notna(x) and x != 'nan' else x)

            # 将'nan'字符串转换回NaN
            df = df.replace('nan', pd.NA)

            print(f"CSV数据: {len(df)} 行, {len(df.columns)} 列")

            # 检查Excel文件类型并写入
            file_extension = Path(excel_file_path).suffix.lower()

            if file_extension == '.xlsx':
                self.write_to_xlsx(df, excel_file_path, sheet_name)
            elif file_extension == '.xls':
                self.write_to_xls(df, excel_file_path, sheet_name)
            else:
                raise ValueError(f"不支持的Excel文件格式: {file_extension}")

            print(f"✅ 成功将CSV数据写入Excel文件: {excel_file_path}")
            print(f"工作表: {sheet_name}")

        except Exception as e:
            raise Exception(f"写入Excel文件失败: {str(e)}")

    def write_to_xlsx(self, df, excel_file_path, sheet_name):
        """将DataFrame写入.xlsx文件的指定工作表"""
        try:
            # 检查文件是否存在
            if Path(excel_file_path).exists():
                # 文件存在，读取现有工作簿
                workbook = openpyxl.load_workbook(excel_file_path)

                # 记录原始工作表顺序
                original_sheet_names = workbook.sheetnames.copy()
                target_sheet_index = -1

                # 如果工作表存在，记录其位置并删除
                if sheet_name in workbook.sheetnames:
                    target_sheet_index = original_sheet_names.index(sheet_name)
                    del workbook[sheet_name]

                # 创建新的工作表
                worksheet = workbook.create_sheet(sheet_name)

                # 如果找到了原始位置，将工作表移动到正确位置
                if target_sheet_index != -1:
                    # 将新工作表移动到原始位置
                    workbook.move_sheet(worksheet, target_sheet_index)

            else:
                # 文件不存在，创建新工作簿
                workbook = openpyxl.Workbook()
                worksheet = workbook.active
                worksheet.title = sheet_name

            # 写入列标题
            for col_idx, column_name in enumerate(df.columns, 1):
                worksheet.cell(row=1, column=col_idx, value=column_name)

            # 写入数据
            for row_idx, (_, row) in enumerate(df.iterrows(), 2):
                for col_idx, value in enumerate(row, 1):
                    # 处理NaN值
                    if pd.isna(value):
                        cell_value = None
                    else:
                        cell_value = value
                    worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

            # 保存文件
            workbook.save(excel_file_path)
            workbook.close()

        except Exception as e:
            raise Exception(f"写入.xlsx文件失败: {str(e)}")

    def write_to_xls(self, df, excel_file_path, sheet_name):
        """将DataFrame写入.xls文件的指定工作表"""
        try:
            # 对于.xls文件，我们需要重新创建整个文件
            # 因为xlwt不支持修改现有文件

            # 如果原文件存在，先读取所有工作表（按原始顺序）
            existing_sheets = []  # 使用列表保持顺序
            target_sheet_index = -1  # 目标工作表的原始位置

            if Path(excel_file_path).exists():
                try:
                    old_workbook = xlrd.open_workbook(excel_file_path)
                    for sheet_idx in range(old_workbook.nsheets):
                        old_sheet = old_workbook.sheet_by_index(sheet_idx)
                        old_sheet_name = old_sheet.name

                        if old_sheet_name == sheet_name:
                            # 记录目标工作表的位置
                            target_sheet_index = sheet_idx
                            # 为目标工作表预留位置
                            existing_sheets.append((old_sheet_name, None))
                        else:
                            # 保存其他工作表的数据
                            sheet_data = []
                            for row_idx in range(old_sheet.nrows):
                                row_data = []
                                for col_idx in range(old_sheet.ncols):
                                    cell_value = old_sheet.cell_value(row_idx, col_idx)
                                    row_data.append(cell_value)
                                sheet_data.append(row_data)
                            existing_sheets.append((old_sheet_name, sheet_data))
                except Exception as e:
                    print(f"警告: 读取原有.xls文件时出错: {str(e)}")

            # 如果没有找到目标工作表，添加到末尾
            if target_sheet_index == -1:
                existing_sheets.append((sheet_name, None))

            # 创建新的工作簿
            new_workbook = xlwt.Workbook()

            # 按原始顺序添加工作表
            for sheet_name_in_order, sheet_data in existing_sheets:
                if sheet_data is None:
                    # 这是目标工作表，写入新数据
                    worksheet = new_workbook.add_sheet(sheet_name_in_order)

                    # 写入列标题
                    for col_idx, column_name in enumerate(df.columns):
                        worksheet.write(0, col_idx, column_name)

                    # 写入数据
                    for row_idx, (_, row) in enumerate(df.iterrows(), 1):
                        for col_idx, value in enumerate(row):
                            # 处理NaN值
                            if pd.isna(value):
                                cell_value = ""
                            else:
                                cell_value = value
                            worksheet.write(row_idx, col_idx, cell_value)
                else:
                    # 这是现有工作表，复制原数据
                    worksheet = new_workbook.add_sheet(sheet_name_in_order)
                    for row_idx, row_data in enumerate(sheet_data):
                        for col_idx, cell_value in enumerate(row_data):
                            worksheet.write(row_idx, col_idx, cell_value)

            # 保存文件
            new_workbook.save(excel_file_path)

        except Exception as e:
            raise Exception(f"写入.xls文件失败: {str(e)}")

    def get_sheet_names(self, excel_file_path):
        """获取Excel文件的工作表名称列表"""
        try:
            file_extension = Path(excel_file_path).suffix.lower()

            if file_extension == '.xlsx':
                workbook = openpyxl.load_workbook(excel_file_path)
                sheet_names = workbook.sheetnames.copy()
                workbook.close()
                return sheet_names
            elif file_extension == '.xls':
                workbook = xlrd.open_workbook(excel_file_path)
                sheet_names = workbook.sheet_names()
                return sheet_names
            else:
                return []
        except Exception as e:
            print(f"警告: 读取工作表名称时出错: {str(e)}")
            return []

    def update_excel_from_csv(self):
        """遍历xls文件夹中的CSV文件，将其内容写回到对应的Excel文件"""
        try:
            print("Excel更新工具")
            print("="*50)
            print(f"CSV文件夹: {self.output_folder}")
            print(f"目标Excel文件夹: {self.target_folder}")
            print("="*50)

            # 查找所有CSV文件
            csv_files = list(self.output_folder.glob("*.csv"))

            if not csv_files:
                print(f"在文件夹 '{self.output_folder}' 中未找到CSV文件")
                return False

            print(f"找到 {len(csv_files)} 个CSV文件:")
            for csv_file in csv_files:
                print(f"  {csv_file.name}")
            print()

            success_count = 0

            # 处理每个CSV文件
            for csv_file in csv_files:
                try:
                    # 解析CSV文件名，提取Excel文件名和工作表名
                    # 格式: filename[sheetname].csv
                    csv_filename = csv_file.stem  # 去掉.csv扩展名

                    if '[' not in csv_filename or ']' not in csv_filename:
                        print(f"⚠️  跳过文件 {csv_file.name}: 文件名格式不正确")
                        continue

                    # 分离文件名和工作表名
                    file_part, sheet_part = csv_filename.split('[', 1)
                    sheet_name = sheet_part.rstrip(']')
                    excel_filename = f"{file_part.strip()}.xls"  # 默认添加.xls扩展名

                    print(f"处理文件: {csv_file.name}")
                    print(f"  目标Excel文件: {excel_filename}")
                    print(f"  目标工作表: {sheet_name}")

                    # 查找对应的Excel文件
                    excel_file_path = self.find_excel_file(excel_filename)
                    if not excel_file_path:
                        print(f"  ❌ 未找到对应的Excel文件: {excel_filename}")
                        continue

                    # 记录更新前的工作表顺序
                    original_sheet_names = self.get_sheet_names(excel_file_path)
                    print(f"  更新前工作表顺序: {original_sheet_names}")

                    # 将CSV内容写入Excel文件
                    self.write_csv_to_excel(csv_file, excel_file_path, sheet_name)

                    # 验证更新后的工作表顺序
                    updated_sheet_names = self.get_sheet_names(excel_file_path)
                    print(f"  更新后工作表顺序: {updated_sheet_names}")

                    # 检查顺序是否保持不变
                    if original_sheet_names == updated_sheet_names:
                        print(f"  ✅ 成功更新，工作表顺序保持不变")
                    else:
                        print(f"  ⚠️  更新成功，但工作表顺序发生变化")

                    success_count += 1

                except Exception as e:
                    print(f"  ❌ 处理失败: {str(e)}")

                print()

            print(f"🎉 更新完成! 成功处理 {success_count}/{len(csv_files)} 个文件")
            return success_count > 0

        except Exception as e:
            print(f"❌ 更新失败: {str(e)}")
            return False

def main():
    """主函数"""
    # 创建转换器实例
    converter = ExcelToCSVConverter(TARGET_FOLDER, OUTPUT_FOLDER)

    # 检查命令行参数
    if len(sys.argv) == 1:
        # 没有参数，执行CSV到Excel的更新操作
        print("Excel更新工具 - 将CSV文件写回Excel")
        print("="*50)
        print(f"目标文件夹: {TARGET_FOLDER}")
        print(f"CSV文件夹: {Path.cwd()}/{OUTPUT_FOLDER}")
        print("="*50)

        success = converter.update_excel_from_csv()

        if success:
            print("\n🎉 更新任务完成!")
        else:
            print("\n💥 更新任务失败!")

    elif len(sys.argv) == 2:
        # 有一个参数，执行Excel到CSV的导出操作
        command = sys.argv[1]

        print("Excel工作表导出为CSV工具")
        print("="*50)
        print(f"目标文件夹: {TARGET_FOLDER}")
        print(f"输出文件夹: {Path.cwd()}/{OUTPUT_FOLDER}")
        print("="*50)
        print(f"执行命令: {command}")
        print()

        success = converter.convert(command)

        if success:
            print("\n🎉 导出任务完成!")
        else:
            print("\n💥 导出任务失败!")
    else:
        # 参数错误
        print("Excel配置工具")
        print("="*50)
        print("使用方法:")
        print("1. 导出Excel工作表为CSV:")
        print("   py config.py filename[sheetname]")
        print("   示例: py config.py hero[hero]")
        print()
        print("2. 将CSV文件写回Excel:")
        print("   py config.py")
        print("   (无参数，自动处理xls文件夹中的所有CSV文件)")
        print("="*50)

if __name__ == "__main__":
    main()
