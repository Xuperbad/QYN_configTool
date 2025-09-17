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

# ==================== 预预处理配置区域 ====================
# 自定义字段关联配置
# 格式: "源表[源工作表], 源列名": "目标表[目标工作表], 匹配列名, 返回列名"
# 示例: "hero[hero], 技能-初始资质": "heroSkill[heroskill], 技能id, 名称"
PRE_PROCESSING_CONFIG = {
    "hero[hero], 技能-初始资质": "heroSkill[heroskill], 技能id, 名称",
    "hero[hero], 技能-商铺": "heroSkill[heroskill], 技能id, 名称",
    "hero[hero], 潜能技能": "heroSkill[heroskill], 技能id, 名称",
    "hero[hero], 光环": "heroSkill[heroskill], 技能id, 名称",
    "wife[wif], 门客缘分": "hero[hero], 人才ID, 名字",
    "wife[wif], 技能": "wife[老婆技能], 序号, 技能名",
    # 可以添加更多关联配置
    # "表名[工作表], 列名": "目标表[目标工作表], 匹配列, 返回列",
}
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
        """查找文本中所有的t_*字符串，包括{}内的t_*字符串"""
        if pd.isna(text) or not isinstance(text, str):
            return []

        t_strings = []

        # 方法1: 查找普通的t_*字符串（按逗号分割）
        parts = [part.strip() for part in text.split(',')]
        for part in parts:
            # 检查每个部分是否是完整的t_*字符串
            if re.match(r'^t_[a-zA-Z0-9_]+$', part):
                t_strings.append(part)

        # 方法2: 查找{}内的t_*字符串，如 1001{t_heroSkillnew_name1001}
        brace_pattern = r'\{(t_[a-zA-Z0-9_]+)\}'
        t_strings_in_braces = re.findall(brace_pattern, text)
        t_strings.extend(t_strings_in_braces)

        # 方法3: 查找所有独立的t_*字符串（不在{}内的）
        # 这个用于处理可能遗漏的情况
        all_t_pattern = r't_[a-zA-Z0-9_]+'
        all_t_strings = re.findall(all_t_pattern, text)

        # 过滤掉已经在{}内的t_*字符串，避免重复处理
        for t_str in all_t_strings:
            # 检查这个t_*字符串是否在{}内
            if f'{{{t_str}}}' not in text:
                t_strings.append(t_str)

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

    def parse_ids_from_value(self, value):
        """从值中解析出ID列表，支持多种格式"""
        if pd.isna(value) or not isinstance(value, str):
            return []

        # 移除空白字符
        value = value.strip()
        if not value:
            return []

        ids = []

        # 处理数组格式 [1001, 1002] 或 [1001,1002]
        if value.startswith('[') and value.endswith(']'):
            # 移除方括号
            inner_value = value[1:-1].strip()
            if inner_value:
                # 按逗号分割
                parts = [part.strip() for part in inner_value.split(',')]
                ids.extend([part for part in parts if part])
        else:
            # 处理逗号分隔格式 1001, 1002 或单个值 1001
            parts = [part.strip() for part in value.split(',')]
            ids.extend([part for part in parts if part])

        return ids

    def pre_preprocess_dataframe(self, df, current_table_sheet):
        """预预处理DataFrame，根据配置进行字段关联查找"""
        if not PRE_PROCESSING_CONFIG:
            print("未配置预预处理规则，跳过预预处理")
            return df

        print("正在进行预预处理，处理自定义字段关联...")

        # 导入go.py模块
        try:
            import sys
            sys.path.append(str(Path.cwd()))
            from go import ExcelTextReplacer
        except Exception as e:
            print(f"导入go.py模块失败: {str(e)}")
            return df

        # 创建查找器实例
        replacer = ExcelTextReplacer({})
        df_processed = df.copy()

        # 处理每个配置规则
        for source_config, target_config in PRE_PROCESSING_CONFIG.items():
            try:
                # 解析源配置: "hero[hero], 技能-初始资质"
                source_parts = [part.strip() for part in source_config.split(',')]
                if len(source_parts) != 2:
                    print(f"源配置格式错误: {source_config}")
                    continue

                source_table_sheet = source_parts[0]  # "hero[hero]"
                source_column = source_parts[1]       # "技能-初始资质"

                # 检查是否匹配当前处理的表和工作表
                if source_table_sheet != current_table_sheet:
                    continue

                # 解析目标配置: "heroSkill[heroskill], 技能id, 名称"
                target_parts = [part.strip() for part in target_config.split(',')]
                if len(target_parts) != 3:
                    print(f"目标配置格式错误: {target_config}")
                    continue

                target_table_sheet = target_parts[0]  # "heroSkill[heroskill]"
                match_column = target_parts[1]        # "技能id"
                return_column = target_parts[2]       # "名称"

                # 解析目标表信息
                if '[' not in target_table_sheet or ']' not in target_table_sheet:
                    print(f"目标表格式错误: {target_table_sheet}")
                    continue

                target_file_part, target_sheet_part = target_table_sheet.split('[', 1)
                target_sheet_name = target_sheet_part.rstrip(']')
                target_filename = f"{target_file_part.strip()}.xls"

                # 构建目标文件的绝对路径
                target_file_path = Path(TARGET_FOLDER) / target_filename

                if not target_file_path.exists():
                    # 尝试其他扩展名
                    for ext in SUPPORTED_EXTENSIONS:
                        test_path = Path(TARGET_FOLDER) / f"{target_file_part.strip()}{ext}"
                        if test_path.exists():
                            target_file_path = test_path
                            break
                    else:
                        print(f"未找到目标文件: {target_filename}")
                        continue

                print(f"处理字段关联: {source_column} -> {target_file_path.name}[{target_sheet_name}]")

                # 检查源列是否存在
                if source_column not in df_processed.columns:
                    print(f"源列 '{source_column}' 不存在")
                    continue

                # 收集所有需要查找的ID
                all_ids = set()
                for idx, cell_value in df_processed[source_column].items():
                    ids = self.parse_ids_from_value(cell_value)
                    all_ids.update(ids)

                if not all_ids:
                    print(f"在列 '{source_column}' 中未找到任何ID")
                    continue

                print(f"找到 {len(all_ids)} 个唯一ID，正在查找对应值...")

                # 使用并发优化的查找方法
                lookup_results = self._lookup_field_values_concurrent(
                    replacer,
                    str(target_file_path),
                    target_sheet_name,
                    match_column,
                    return_column,
                    list(all_ids)
                )

                found_count = len(lookup_results)
                print(f"成功找到 {found_count}/{len(all_ids)} 个ID的对应值")

                # 替换DataFrame中的内容
                for idx, cell_value in df_processed[source_column].items():
                    if pd.notna(cell_value) and isinstance(cell_value, str):
                        ids = self.parse_ids_from_value(cell_value)
                        if ids:
                            # 构建新值，格式: id{对应值}
                            new_parts = []
                            for id_val in ids:
                                if id_val in lookup_results:
                                    new_parts.append(f"{id_val}{{{lookup_results[id_val]}}}")
                                else:
                                    new_parts.append(id_val)  # 保持原值

                            # 根据原格式重新组装
                            original_value = str(cell_value).strip()
                            if original_value.startswith('[') and original_value.endswith(']'):
                                # 保持数组格式
                                df_processed.loc[idx, source_column] = f"[{', '.join(new_parts)}]"
                            else:
                                # 保持逗号分隔格式
                                df_processed.loc[idx, source_column] = ', '.join(new_parts)

                print(f"完成字段 '{source_column}' 的关联处理")

            except Exception as e:
                print(f"处理配置 '{source_config}' 时出错: {str(e)}")
                continue

        print("预预处理完成")
        return df_processed

    def _lookup_field_values_concurrent(self, replacer, file_path, sheet_name, match_column, return_column, search_values):
        """并发优化的字段值查找方法"""
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import multiprocessing

        if not search_values:
            return {}

        # 动态调整线程数
        cpu_count = multiprocessing.cpu_count()
        max_workers = min(max(cpu_count * 2, 4), len(search_values), 16)  # 预预处理用较少线程

        print(f"使用 {max_workers} 个线程并发查找字段值...")

        # 将搜索值分批处理
        batch_size = max(1, len(search_values) // max_workers)
        batches = [search_values[i:i + batch_size] for i in range(0, len(search_values), batch_size)]

        def lookup_batch(batch_values):
            """查找一批值"""
            return replacer.lookup_field_values(file_path, sheet_name, match_column, return_column, batch_values)

        # 并发执行查找
        all_results = {}
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 提交所有批次任务
            future_to_batch = {executor.submit(lookup_batch, batch): batch for batch in batches}

            # 收集结果
            completed_count = 0
            for future in as_completed(future_to_batch):
                batch = future_to_batch[future]
                try:
                    batch_results = future.result()
                    all_results.update(batch_results)
                    completed_count += len(batch)
                    print(f"  已完成 {completed_count}/{len(search_values)} 个ID的查找")
                except Exception as e:
                    print(f"  批次查找失败: {str(e)}")

        return all_results

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

                    # 方法1: 直接替换完整的t_*字符串（按逗号分割）
                    parts = [part.strip() for part in new_value.split(',')]
                    new_parts = []

                    for part in parts:
                        # 检查这个部分是否是完整的t_*字符串
                        if part in t_string_map:
                            new_parts.append(t_string_map[part])
                        else:
                            new_parts.append(part)

                    new_value = ', '.join(new_parts)

                    # 方法2: 替换{}内的t_*字符串
                    import re
                    def replace_t_in_braces(match):
                        t_string = match.group(1)  # 获取{}内的t_*字符串
                        if t_string in t_string_map:
                            return '{' + t_string_map[t_string] + '}'
                        else:
                            return match.group(0)  # 保持原样

                    # 使用正则表达式替换{}内的t_*字符串
                    new_value = re.sub(r'\{(t_[a-zA-Z0-9_]+)\}', replace_t_in_braces, new_value)

                    df_processed.loc[idx, col] = new_value

        print(f"预处理完成，共找到 {found_count}/{len(all_t_strings)} 个t_*字符串的中文对应")
        print(f"替换了 {found_count} 个t_*字符串为带中文的格式")
        return df_processed

    def save_to_csv(self, dataframe, output_filename):
        """将DataFrame保存为CSV文件"""
        output_path = self.output_folder / output_filename

        try:
            # 保存为CSV，使用UTF-8编码，禁用引号转义
            dataframe.to_csv(output_path, index=False, encoding='utf-8-sig',
                           quoting=1, escapechar=None)  # quoting=1 表示 QUOTE_ALL
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

            # 构建当前表和工作表的标识
            base_filename = Path(filename).stem  # 去掉扩展名
            current_table_sheet = f"{base_filename}[{sheet_name}]"

            # 预预处理DataFrame，处理自定义字段关联
            df_pre_processed = self.pre_preprocess_dataframe(df, current_table_sheet)

            # 预处理DataFrame，查找并替换t_*字符串
            df_processed = self.preprocess_dataframe(df_pre_processed)

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
        """处理CSV内容，将各种{中文}格式还原为原始格式"""
        import re

        # 正则表达式匹配 t_*{中文} 格式
        t_pattern = r't_([a-zA-Z0-9_]+)\{[^}]*\}'
        # 替换为 t_* 格式
        processed_content = re.sub(t_pattern, r't_\1', csv_content)

        # 正则表达式匹配 数字{中文} 格式（用于ID关联）
        id_pattern = r'(\d+)\{[^}]*\}'
        # 替换为纯数字格式
        processed_content = re.sub(id_pattern, r'\1', processed_content)

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
