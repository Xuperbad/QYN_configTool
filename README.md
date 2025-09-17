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
- **预预处理**: 自定义字段关联，支持跨表数据查找
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

#### 1. 配置预预处理规则（新功能）
打开 `config.py` 文件，找到预预处理配置区域：

```python
# ==================== 预预处理配置区域 ====================
# 自定义字段关联配置
# 格式: "源表[源工作表], 源列名": "目标表[目标工作表], 匹配列名, 返回列名"
PRE_PROCESSING_CONFIG = {
    "hero[hero], 技能-初始资质": "heroSkill[heroskill], 技能id, 名称",
    "hero[hero], 技能-商铺": "heroSkill[heroskill], 技能id, 名称",
    "hero[hero], 潜能技能": "heroSkill[heroskill], 技能id, 名称",
    # 可以添加更多关联配置
    # "表名[工作表], 列名": "目标表[目标工作表], 匹配列, 返回列",
}
```

**配置说明**:
- **源表配置**: `"hero[hero], 技能-初始资质"` 表示处理 `hero.xls` 文件的 `hero` 工作表中的 `技能-初始资质` 列
- **目标表配置**: `"heroSkill[heroskill], 技能id, 名称"` 表示在 `heroSkill.xls` 文件的 `heroskill` 工作表中，用 `技能id` 列匹配，返回 `名称` 列的值
- **处理效果**: 将 `[1001, 1014, 1051]` 转换为 `[1001{t_heroSkillnew_name1001}, 1014{t_heroSkillnew_name1014}, 1051{t_heroSkillnew_name1051}]`

**支持的数据格式**:
- 单个值: `1001`
- 逗号分隔: `1001, 1002`
- 数组格式: `[1001, 1002]`

#### 2. 导出Excel工作表为CSV
```bash
# 导出指定Excel文件的指定工作表为CSV
python config.py filename[sheetname]

# 示例：导出hero.xls文件的hero工作表
python config.py hero[hero]
```

这个命令会：
- 读取 `E:\qyn_game\parseFiles\global\config\test\hero.xls` 文件的 `hero` 工作表
- **预预处理**: 根据配置进行字段关联查找，添加关联数据的注释
- **预处理**: 自动识别并处理 `t_*` 格式的ID字符串，添加中文翻译
- 将结果保存为 `xls/hero[hero].csv`

#### 3. 将修改后的CSV写回Excel（新功能）
```bash
# 将xls文件夹中的所有CSV文件写回对应的Excel文件
python config.py
```

这个命令会：
- 扫描 `xls` 文件夹中的所有CSV文件
- 解析文件名格式 `filename[sheetname].csv`
- 将 `t_*{中文}` 和 `id{关联值}` 格式还原为原始格式
- 将数据写回到对应的Excel文件和工作表中
- **保持工作表顺序不变**

#### 完整工作流程
1. **导出**: `python config.py hero[hero]` - 将Excel导出为CSV，带中文注释和关联数据注释
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

## 功能特性

### 预预处理功能
- **跨表数据关联**: 支持从其他Excel文件查找关联数据
- **智能格式识别**: 自动识别单值、逗号分隔、数组等格式
- **批量处理**: 一次性处理多个字段的关联查找
- **性能优化**: 批量查找，避免重复读取文件

### 智能并发处理
- **动态线程数**: 根据CPU核心数和任务量自动调整
- **高效搜索**: 32线程并发搜索t_*字符串对应的中文
- **进度显示**: 实时显示搜索进度和结果

### 数据完整性
- **工作表顺序保持**: 更新Excel文件时保持原有工作表顺序
- **格式还原**: CSV写回时智能还原各种格式
- **错误处理**: 完善的错误提示和异常处理

## 输出说明

### CSV文件格式
- CSV文件保存在 `xls` 文件夹中
- 文件名格式：`filename[sheetname].csv`
- 包含中文注释和关联数据注释，便于编辑和理解

### Excel文件更新
- 原Excel文件会被直接更新
- 保持原有的文件结构和其他工作表不变
- 只更新指定的工作表内容

## 注意事项

1. **文件路径**：确保配置的目标文件夹路径正确
2. **文件格式**：支持 .xlsx 和 .xls 格式
3. **备份建议**：在批量处理前建议备份重要文件
4. **编码格式**：CSV文件使用UTF-8编码，支持中文
5. **关联表格**：确保关联的目标表格存在且格式正确

## 故障排除

### 常见问题：
1. **找不到Excel文件**
   - 检查 `TARGET_FOLDER` 配置是否正确
   - 确认文件名和路径是否存在

2. **关联查找失败**
   - 检查目标表格是否存在
   - 确认列名是否正确
   - 验证数据格式是否匹配

3. **CSV文件格式错误**
   - 确保CSV文件名格式为 `filename[sheetname].csv`
   - 检查CSV文件编码是否为UTF-8

4. **权限错误**
   - 确保对目标目录有读写权限
   - 确保Excel文件没有被其他程序占用

5. **依赖包缺失**
   ```bash
   pip install pandas openpyxl xlrd xlwt
   ```

## 示例

### 预预处理示例
**原始数据**: `[1001, 1014, 1051]`
**处理后**: `[1001{t_heroSkillnew_name1001}, 1014{t_heroSkillnew_name1014}, 1051{t_heroSkillnew_name1051}]`

### t_*字符串处理示例
**原始数据**: `t_heronew_name500001`
**处理后**: `t_heronew_name500001{五竹}`

这样您就可以在CSV文件中直观地看到每个ID对应的实际含义，大大提高了数据的可读性和编辑效率！
