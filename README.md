# QYN配置工具

这是一个用于处理Excel配置文件的工具集，包含文本替换和CSV导出功能，专门用于游戏配置文件的处理。

## 功能介绍

### 1. Excel文本替换工具 (go.py)
- 批量替换Excel文件中的文本内容
- 支持精确搜索和模糊搜索
- 保持原有格式和样式不变
- 支持.xlsx和.xls格式

### 2. Excel转CSV工具 (config.py)
- 将Excel工作表导出为CSV格式
- 自动识别和处理t_*格式的ID字符串
- 并发查找对应的中文翻译
- 智能处理唯一性检查

## 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法

### Excel文本替换工具 (go.py)

#### 1. 配置替换规则
打开 `go.py` 文件，找到配置区域：

```python
# ==================== 配置区域 ====================
REPLACEMENT_CONFIG = {
    "人才":"能士",
    "知己":"挚友",
    "士人":"势者",
    "农民":"力者",
    "工匠":"韧者",
    "商贾":"智者",
    "武者":"敏者",
    # 可以添加更多替换规则，格式为 "原文本": "新文本"
}
```

#### 2. 运行替换操作
```bash
# 处理当前目录下的所有Excel文件
python go.py

# 处理指定目录下的所有Excel文件
python go.py /path/to/excel/files

# 搜索特定文本
python go.py "t_heronew_name500001"

# 模糊搜索（前缀匹配）
python go.py "t_hero*"
```

### Excel转CSV工具 (config.py)

#### 基本用法
```bash
# 导出指定工作表为CSV
python config.py filename[sheetname]

# 示例：导出hero.xls文件中的hero工作表
python config.py hero[hero]
```

#### 功能特点
- 自动识别t_*格式的ID字符串
- 并发查找对应的中文翻译（基于CPU核心数自动调整线程数）
- 智能处理：
  - 唯一结果：`t_heronew_name500001{五竹}`
  - 非唯一结果：保持原样 `t_heronew_name500001`
  - 未找到：保持原样

## 配置说明

### 目标文件夹配置
在 `go.py` 和 `config.py` 中都有 `TARGET_FOLDER` 配置：

```python
# 目标文件夹路径
TARGET_FOLDER = r"E:\qyn_game\parseFiles\global\config\test"
```

根据实际情况修改此路径。

### 输出文件夹配置 (config.py)
```python
# 输出文件夹名称（在当前工作目录下创建）
OUTPUT_FOLDER = "xls"
```

## 输出说明

### go.py 输出
- **替换模式**：直接修改原文件
- **搜索模式**：显示搜索结果，格式为：
  ```
  tableLang.xls[functionLang], 行1021: t_heronew_name500001, 五竹
  ```

### config.py 输出
- **CSV文件**：保存在当前目录的 `xls` 文件夹下
- **文件名格式**：`原文件名[工作表名].csv`
- **处理过程**：显示并发搜索进度和替换统计

## 使用示例

### 文本替换示例
假设Excel文件包含：
- "人才管理系统"
- "优秀人才培养"

配置替换规则为 `"人才": "能士"` 后，运行 `python go.py`，结果变为：
- "能士管理系统"
- "优秀能士培养"

### CSV导出示例
```bash
# 导出hero.xls文件中的hero工作表
python config.py hero[hero]
```

处理前的数据：
```
ID,Name,Description
1,t_heronew_name500001,这是一个角色
```

处理后的数据：
```
ID,Name,Description
1,t_heronew_name500001{五竹},这是一个角色
```

## 性能优化

### 并发处理
- `config.py` 支持并发处理，自动根据CPU核心数调整线程数
- 默认使用 `CPU核心数 × 3` 个线程，最多32个线程
- 对于大量t_*字符串的处理，可显著提升速度

### 内存优化
- 直接读取Excel文件，避免字符串解析
- 使用高效的数据结构存储搜索结果

## 注意事项

1. **备份重要文件**：go.py会直接修改原文件，建议先备份
2. **文件格式**：支持 .xlsx 和 .xls 格式
3. **样式保持**：go.py会保持原有的字体、颜色、边框等格式
4. **大小写敏感**：替换和搜索都是大小写敏感的
5. **唯一性检查**：config.py会检查t_*字符串的唯一性，非唯一结果不会被处理

## 故障排除

### 常见错误：
1. **ModuleNotFoundError**
   ```bash
   pip install openpyxl xlrd pandas
   ```

2. **文件被占用**
   - 确保Excel文件没有在其他程序中打开

3. **路径错误**
   - 检查 `TARGET_FOLDER` 配置是否正确
   - 确保目标文件存在

4. **编码问题**
   - 确保Excel文件编码正确
   - 中文内容应保存为UTF-8格式

### 调试技巧：
- 使用搜索模式测试：`python go.py "your_search_term"`
- 查看详细的控制台输出
- 检查生成的CSV文件内容
