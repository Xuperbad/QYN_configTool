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
- **新功能**: 将修改后的CSV文件写回Excel文件

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

#### 2. 运行替换
```bash
# 替换当前目录下的所有Excel文件
python go.py

# 替换指定目录下的所有Excel文件
python go.py /path/to/excel/files

# 搜索文本（精确搜索）
python go.py "搜索文本"

# 模糊搜索（前缀匹配）
python go.py "t_hero_getway*"
```

### Excel转CSV工具 (config.py)

#### 1. 导出Excel工作表为CSV
```bash
# 导出指定Excel文件的指定工作表为CSV
python config.py filename[sheetname]

# 示例：导出hero.xls文件的hero工作表
python config.py hero[hero]
```

这个命令会：
- 读取 `E:\qyn_game\parseFiles\global\config\test\hero.xls` 文件的 `hero` 工作表
- 自动识别并处理 `t_*` 格式的ID字符串
- 并发查找对应的中文翻译，将 `t_heronew_name500001` 转换为 `t_heronew_name500001{五竹}`
- 将结果保存为 `xls/hero[hero].csv`

#### 2. 将修改后的CSV写回Excel（新功能）
```bash
# 将xls文件夹中的所有CSV文件写回对应的Excel文件
python config.py
```

这个命令会：
- 扫描 `xls` 文件夹中的所有CSV文件
- 解析文件名格式 `filename[sheetname].csv`
- 将 `t_*{中文}` 格式还原为 `t_*` 格式
- 将数据写回到对应的Excel文件和工作表中

#### 完整工作流程
1. **导出**: `python config.py hero[hero]` - 将Excel导出为CSV，带中文注释
2. **编辑**: 直接编辑 `xls/hero[hero].csv` 文件，修改或添加内容
3. **写回**: `python config.py` - 将修改后的CSV写回Excel文件

## 配置说明

### 目标文件夹配置
在 `config.py` 和 `go.py` 中都有目标文件夹配置：

```python
# config.py 中的配置
TARGET_FOLDER = r"E:\qyn_game\parseFiles\global\config\test"

# go.py 中的配置
TARGET_FOLDER = r"E:\qyn_game\parseFiles\global\config\test\lang_client"
```

根据你的实际路径修改这些配置。

## 输出说明

### CSV文件格式
- CSV文件保存在 `xls` 文件夹中
- 文件名格式：`filename[sheetname].csv`
- 包含中文注释的t_*字符串，便于编辑和理解

### Excel文件更新
- 原Excel文件会被直接更新
- 保持原有的文件结构和其他工作表不变
- 只更新指定的工作表内容

## 注意事项

1. **文件路径**：确保配置的目标文件夹路径正确
2. **文件格式**：支持 .xlsx 和 .xls 格式
3. **备份建议**：在批量处理前建议备份重要文件
4. **编码格式**：CSV文件使用UTF-8编码，支持中文
5. **并发处理**：t_*字符串查找使用多线程，提高处理速度

## 故障排除

### 常见问题：
1. **找不到Excel文件**
   - 检查 `TARGET_FOLDER` 配置是否正确
   - 确认文件名和路径是否存在

2. **CSV文件格式错误**
   - 确保CSV文件名格式为 `filename[sheetname].csv`
   - 检查CSV文件编码是否为UTF-8

3. **权限错误**
   - 确保对目标目录有读写权限
   - 确保Excel文件没有被其他程序占用

4. **依赖包缺失**
   ```bash
   pip install pandas openpyxl xlrd xlwt
   ```
