# Python Excel 交互操作项目

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)]()

这是一个功能完整的 Python Excel 交互操作项目，使用 `win32com.client` 库实现与 Microsoft Excel 的深度集成。项目提供了丰富的 Excel 自动化功能，从基础的数据操作到高级的图表创建和宏执行。

## ✨ 功能特性

### 🔧 核心功能
- **Excel 基础操作** - 创建、读取、写入、保存 Excel 文件
- **工作表管理** - 新建、删除、重命名、复制工作表
- **数据处理** - 智能数据读写、查找替换、批量操作
- **格式设置** - 字体、颜色、边框、对齐、数字格式

### 📊 高级功能
- **图表操作** - 创建柱状图、折线图、饼图等多种图表类型
- **数据透视表** - 创建、配置、刷新数据透视表
- **宏操作** - 运行 VBA 宏、创建 VBA 模块
- **打印功能** - 页面设置、打印预览、PDF 导出
- **外部数据** - 数据库连接、外部数据源刷新

### 🛡️ 企业级特性
- **异常处理** - 完善的错误处理和恢复机制
- **性能优化** - 批量操作、屏幕更新控制
- **日志记录** - 详细的操作日志和调试信息
- **模块化设计** - 清晰的代码结构，易于维护和扩展

## 🏗️ 项目结构

```
excel_interaction_project/
├── 📄 main.py                    # 主程序入口，演示所有功能
├── ⚙️ config.py                  # 配置文件
├── 📋 requirements.txt           # 项目依赖
├── 📁 modules/                   # 功能模块目录
│   ├── 📝 excel_basic.py        # Excel基础操作
│   ├── 🔧 excel_macro.py        # 宏操作和VBA调用
│   ├── 📊 excel_pivot.py        # 数据透视表操作
│   ├── 📈 excel_chart.py        # 图表创建和操作
│   ├── 💾 excel_data.py         # 数据处理和连接
│   ├── 🖨️ excel_print.py        # 打印设置和操作
│   └── 🎨 excel_format.py       # 格式设置和样式
├── 📁 utils/                     # 工具模块目录
│   ├── 🎛️ excel_manager.py      # Excel应用程序管理器
│   └── 📋 constants.py          # 常量和配置定义
├── 📁 examples/                  # 示例代码目录
│   ├── 📝 basic_example.py      # 基础操作示例
│   └── 📊 chart_example.py      # 图表操作示例
├── 📁 output/                    # 输出文件目录
└── 📁 logs/                      # 日志文件目录
```

## 🚀 快速开始

### 系统要求

- **操作系统**: Windows 7/8/10/11
- **Python**: 3.7 或更高版本
- **Microsoft Excel**: 2010 或更高版本
- **内存**: 建议 4GB 以上

### 安装步骤

1. **克隆项目**
   ```bash
   git clone https://github.com/your-username/python-excel-automation.git
   cd python-excel-automation
   ```

2. **创建虚拟环境**（推荐）
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   ```

3. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

4. **运行演示程序**
   ```bash
   python main.py
   ```

### 快速示例

```python
from utils.excel_manager import ExcelManager
from modules.excel_basic import ExcelBasic

# 使用上下文管理器确保资源正确释放
with ExcelManager(visible=True) as excel:
    basic = ExcelBasic(excel)
    
    # 创建新工作簿
    wb = basic.create_workbook()
    ws = wb.ActiveSheet
    
    # 写入数据
    data = [['姓名', '年龄', '城市'],
            ['张三', 25, '北京'],
            ['李四', 30, '上海']]
    basic.set_range_values(ws, 'A1:C3', data)
    
    # 保存文件
    basic.save_workbook(wb, 'example.xlsx')
```

## 📖 详细使用说明

### Excel 基础操作

```python
from modules.excel_basic import ExcelBasic

# 创建Excel基础操作实例
basic = ExcelBasic(excel_manager)

# 创建工作簿
workbook = basic.create_workbook()

# 添加工作表
worksheet = basic.add_worksheet(workbook, '数据表')

# 写入数据
basic.set_cell_value(worksheet, 1, 1, '标题')
basic.set_range_values(worksheet, 'A2:C4', data)

# 查找和替换
found_cells = basic.find_cells(worksheet, '查找内容')
basic.replace_cells(worksheet, '旧值', '新值')
```

### 图表创建

```python
from modules.excel_chart import ExcelChart

chart_module = ExcelChart(excel_manager)

# 创建柱状图
chart = chart_module.create_chart(
    worksheet, 'A1:B5', 'xlColumnClustered', 'D2'
)

# 设置图表标题
chart_module.set_chart_title(chart, '销售数据图表')

# 设置轴标题
chart_module.set_axis_title(chart, 'x', '月份')
chart_module.set_axis_title(chart, 'y', '销售额')
```

### 数据透视表

```python
from modules.excel_pivot import ExcelPivot

pivot_module = ExcelPivot(excel_manager)

# 创建数据透视表
pivot_table = pivot_module.create_pivot_table(
    source_worksheet, 'A1:D100', 
    target_worksheet, 'F2', 
    '销售数据透视表'
)

# 添加字段
pivot_module.add_row_field(pivot_table, '产品')
pivot_module.add_column_field(pivot_table, '地区')
pivot_module.add_data_field(pivot_table, '销售额')
```

## 🎯 应用场景

### 📊 数据分析和报告
- 自动生成销售报告
- 财务数据分析
- 业务指标监控
- 数据可视化

### 🔄 数据处理自动化
- 批量数据导入导出
- 数据清洗和转换
- 多文件数据合并
- 定期报表生成

### 📈 商业智能
- 动态仪表板创建
- 趋势分析图表
- 数据透视表分析
- KPI 监控报告

## ⚠️ 注意事项

### 系统要求
- ✅ 必须在 Windows 系统上运行
- ✅ 需要安装 Microsoft Excel（2010或更高版本）
- ✅ 确保 Excel 宏安全设置允许运行宏
- ✅ 建议在虚拟环境中运行项目

### 性能建议
- 🚀 处理大量数据时建议关闭 Excel 界面显示
- 🚀 使用批量操作而非逐个单元格操作
- 🚀 及时释放 COM 对象避免内存泄漏

### 安全提醒
- 🔒 运行宏功能时请确保文件来源可信
- 🔒 建议在测试环境中先验证代码
- 🔒 重要数据请提前备份

## 🤝 贡献指南

我们欢迎所有形式的贡献！

1. **Fork** 本项目
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 **Pull Request**

### 开发环境设置

```bash
# 克隆项目
git clone https://github.com/your-username/python-excel-automation.git
cd python-excel-automation

# 安装开发依赖
pip install -r requirements.txt
pip install pytest black mypy

# 运行测试
pytest

# 代码格式化
black .

# 类型检查
mypy .
```

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

- 感谢 Microsoft 提供强大的 Excel COM 接口
- 感谢 Python 社区的 pywin32 项目
- 感谢所有贡献者和用户的支持

## 📞 联系我们

- 📧 邮箱: hkz638@163.com
- 💬 QQ: 277915799

---

⭐ 如果这个项目对您有帮助，请给我们一个星标！