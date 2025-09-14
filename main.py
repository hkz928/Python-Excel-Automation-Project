# -*- coding: utf-8 -*-
"""
@Time    ：2025/9/14 上午11:44
@Author  ：庄洪奎（ARTHUR)
@FileName：main.py
@Software：PyCharm
@Subions ：zhk0459
"""
"""
Excel 自动化操作主程序

这是一个完整的 Python Excel 交互操作项目的主程序文件。
展示了如何使用各个模块来实现 Excel 的自动化操作。

功能演示：
1. Excel 基础操作
2. 宏操作和 VBA 调用
3. 数据透视表操作
4. 图表操作
5. 外部数据连接
6. 打印设置
7. 格式设置
"""

import os
import sys
import time
from pathlib import Path
from loguru import logger

# 添加项目根目录到路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# 导入模块
from utils.excel_manager import ExcelManager
from modules.excel_basic import ExcelBasic
from modules.excel_macro import ExcelMacro
from modules.excel_pivot import ExcelPivot
from modules.excel_chart import ExcelChart
from modules.excel_data import ExcelData
from modules.excel_print import ExcelPrint
from modules.excel_format import ExcelFormat
from config import get_config, OUTPUT_DIR

# 配置日志
log_config = get_config('log')
logger.add(
    log_config['log_file'],
    format=log_config['format'],
    level=log_config['level'],
    rotation=log_config['rotation'],
    retention=log_config['retention']
)


def demo_basic_operations(excel_mgr: ExcelManager) -> None:
    """
    演示基础操作
    
    Args:
        excel_mgr: Excel 管理器
    """
    logger.info("=== 开始演示基础操作 ===")
    
    try:
        basic = ExcelBasic(excel_mgr)
        
        # 创建工作簿
        wb = basic.create_workbook()
        ws = wb.ActiveSheet
        ws.Name = "基础操作演示"
        
        # 准备示例数据
        data = [
            ['员工姓名', '部门', '职位', '工资', '入职日期'],
            ['张三', '技术部', '工程师', 8000, '2023-01-15'],
            ['李四', '销售部', '销售经理', 12000, '2022-06-20'],
            ['王五', '技术部', '高级工程师', 15000, '2021-03-10'],
            ['赵六', '人事部', '人事专员', 6000, '2023-05-08'],
            ['钱七', '财务部', '会计', 7000, '2022-11-30']
        ]
        
        # 写入数据
        basic.set_range_values(ws, 'A1:E6', data)
        
        # 查找数据
        found_cells = basic.find_cells(ws, '技术部')
        logger.info(f"找到'技术部'的位置: {found_cells}")
        
        # 替换数据
        replaced_count = basic.replace_cells(ws, '技术部', 'IT部')
        logger.info(f"替换了 {replaced_count} 个单元格")
        
        # 自动调整列宽
        basic.auto_fit_columns(ws)
        
        # 保存文件
        output_file = OUTPUT_DIR / "basic_operations_demo.xlsx"
        basic.save_workbook(wb, str(output_file))
        
        logger.info(f"基础操作演示完成，文件保存至: {output_file}")
        
    except Exception as e:
        logger.error(f"基础操作演示失败: {e}")
        raise


def demo_macro_operations(excel_mgr: ExcelManager) -> None:
    """
    演示宏操作
    
    Args:
        excel_mgr: Excel 管理器
    """
    logger.info("=== 开始演示宏操作 ===")
    
    try:
        basic = ExcelBasic(excel_mgr)
        macro = ExcelMacro(excel_mgr)
        
        # 创建工作簿
        wb = basic.create_workbook()
        ws = wb.ActiveSheet
        ws.Name = "宏操作演示"
        
        # 启用宏
        macro.enable_macros(wb)
        
        # 创建 VBA 代码
        vba_code = """
Public Function CalculateBonus(salary As Double, performance As String) As Double
    Select Case performance
        Case "优秀"
            CalculateBonus = salary * 0.2
        Case "良好"
            CalculateBonus = salary * 0.1
        Case "一般"
            CalculateBonus = salary * 0.05
        Case Else
            CalculateBonus = 0
    End Select
End Function

Sub FormatSalaryData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 设置标题格式
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 设置数据格式
    ws.Columns("D:D").NumberFormat = "#,##0"
    ws.Columns("E:E").NumberFormat = "yyyy-mm-dd"
    
    ' 自动调整列宽
    ws.Columns.AutoFit
    
    MsgBox "工资数据格式化完成！"
End Sub

Sub HighlightHighSalary()
    Dim ws As Worksheet
    Dim cell As Range
    Set ws = ActiveSheet
    
    ' 高亮显示高工资（>10000）
    For Each cell In ws.Range("D2:D100")
        If IsNumeric(cell.Value) And cell.Value > 10000 Then
            cell.Interior.Color = RGB(255, 255, 0)  ' 黄色背景
            cell.Font.Bold = True
        End If
    Next cell
    
    MsgBox "高工资数据已高亮显示！"
End Sub
"""
        
        # 创建 VBA 模块
        macro.create_vba_module(wb, "SalaryUtils", vba_code)
        
        # 准备数据
        data = [
            ['员工姓名', '部门', '职位', '工资', '绩效'],
            ['张三', 'IT部', '工程师', 8000, '良好'],
            ['李四', '销售部', '销售经理', 12000, '优秀'],
            ['王五', 'IT部', '高级工程师', 15000, '优秀'],
            ['赵六', '人事部', '人事专员', 6000, '一般'],
            ['钱七', '财务部', '会计', 7000, '良好']
        ]
        
        basic.set_range_values(ws, 'A1:E6', data)
        
        # 运行格式化宏
        try:
            macro.run_macro("SalaryUtils.FormatSalaryData", workbook=wb)
            logger.info("格式化宏执行成功")
        except Exception as e:
            logger.warning(f"运行格式化宏失败: {e}")
        
        # 运行高亮宏
        try:
            macro.run_macro("SalaryUtils.HighlightHighSalary", workbook=wb)
            logger.info("高亮宏执行成功")
        except Exception as e:
            logger.warning(f"运行高亮宏失败: {e}")
        
        # 获取宏列表
        macros = macro.get_macro_list(wb)
        logger.info(f"工作簿中的宏: {macros}")
        
        # 保存文件
        output_file = OUTPUT_DIR / "macro_operations_demo.xlsm"
        basic.save_workbook(wb, str(output_file), file_format='xlOpenXMLWorkbookMacroEnabled')
        
        logger.info(f"宏操作演示完成，文件保存至: {output_file}")
        
    except Exception as e:
        logger.error(f"宏操作演示失败: {e}")
        raise


def demo_pivot_operations(excel_mgr: ExcelManager) -> None:
    """
    演示数据透视表操作
    
    Args:
        excel_mgr: Excel 管理器
    """
    logger.info("=== 开始演示数据透视表操作 ===")
    
    try:
        basic = ExcelBasic(excel_mgr)
        pivot = ExcelPivot(excel_mgr)
        
        # 创建工作簿
        wb = basic.create_workbook()
        ws = wb.ActiveSheet
        ws.Name = "销售数据"
        
        # 准备销售数据
        sales_data = [
            ['日期', '产品', '地区', '销售员', '数量', '单价', '总额'],
            ['2023-01-01', '产品A', '北京', '张三', 10, 100, 1000],
            ['2023-01-01', '产品B', '上海', '李四', 5, 200, 1000],
            ['2023-01-02', '产品A', '广州', '王五', 8, 100, 800],
            ['2023-01-02', '产品C', '北京', '张三', 12, 150, 1800],
            ['2023-01-03', '产品B', '上海', '李四', 6, 200, 1200],
            ['2023-01-03', '产品A', '深圳', '赵六', 15, 100, 1500],
            ['2023-01-04', '产品C', '广州', '王五', 9, 150, 1350],
            ['2023-01-04', '产品B', '北京', '张三', 7, 200, 1400],
            ['2023-01-05', '产品A', '上海', '李四', 11, 100, 1100],
            ['2023-01-05', '产品C', '深圳', '赵六', 13, 150, 1950]
        ]
        
        basic.set_range_values(ws, 'A1:G11', sales_data)
        
        # 创建数据透视表
        pivot_table = pivot.create_pivot_table(
            source_worksheet=ws,
            source_range='A1:G11',
            target_worksheet=ws,
            target_cell='I1',
            table_name='销售数据透视表'
        )
        
        # 配置透视表字段
        pivot.add_row_field(pivot_table, '产品')
        pivot.add_column_field(pivot_table, '地区')
        pivot.add_data_field(pivot_table, '总额', 'xlSum', '销售总额')
        pivot.add_data_field(pivot_table, '数量', 'xlSum', '销售数量')
        
        # 格式化透视表
        pivot.format_pivot_table(pivot_table, 'TableStyleMedium9')
        
        # 创建第二个透视表（按销售员统计）
        pivot_table2 = pivot.create_pivot_table(
            source_worksheet=ws,
            source_range='A1:G11',
            target_worksheet=ws,
            target_cell='I15',
            table_name='销售员业绩透视表'
        )
        
        pivot.add_row_field(pivot_table2, '销售员')
        pivot.add_data_field(pivot_table2, '总额', 'xlSum', '销售总额')
        pivot.add_data_field(pivot_table2, '数量', 'xlSum', '销售数量')
        pivot.add_data_field(pivot_table2, '总额', 'xlAverage', '平均销售额')
        
        # 格式化第二个透视表
        pivot.format_pivot_table(pivot_table2, 'TableStyleMedium6')
        
        # 获取透视表列表
        pivot_tables = pivot.get_pivot_table_list(wb)
        logger.info(f"创建的数据透视表: {len(pivot_tables)} 个")
        
        # 保存文件
        output_file = OUTPUT_DIR / "pivot_operations_demo.xlsx"
        basic.save_workbook(wb, str(output_file))
        
        logger.info(f"数据透视表演示完成，文件保存至: {output_file}")
        
    except Exception as e:
        logger.error(f"数据透视表演示失败: {e}")
        raise


def demo_chart_operations(excel_mgr: ExcelManager) -> None:
    """
    演示图表操作
    
    Args:
        excel_mgr: Excel 管理器
    """
    logger.info("=== 开始演示图表操作 ===")
    
    try:
        basic = ExcelBasic(excel_mgr)
        chart_mgr = ExcelChart(excel_mgr)
        
        # 创建工作簿
        wb = basic.create_workbook()
        ws = wb.ActiveSheet
        ws.Name = "图表演示"
        
        # 准备月度销售数据
        monthly_data = [
            ['月份', '产品A', '产品B', '产品C'],
            ['1月', 1200, 800, 600],
            ['2月', 1400, 900, 700],
            ['3月', 1100, 1200, 800],
            ['4月', 1600, 1000, 900],
            ['5月', 1800, 1100, 1000],
            ['6月', 2000, 1300, 1200]
        ]
        
        basic.set_range_values(ws, 'A1:D7', monthly_data)
        
        # 创建柱状图
        column_chart = chart_mgr.create_chart(
            worksheet=ws,
            data_range='A1:D7',
            chart_type='xlColumnClustered',
            position='F2',
            width=400,
            height=300
        )
        
        # 设置图表标题和坐标轴
        chart_mgr.set_chart_title(column_chart, '月度产品销售对比')
        chart_mgr.set_axis_title(column_chart, 'x', '月份')
        chart_mgr.set_axis_title(column_chart, 'y', '销售额')
        chart_mgr.set_legend(column_chart, 'bottom')
        chart_mgr.set_data_labels(column_chart, show_value=True)
        
        # 创建折线图
        line_chart = chart_mgr.create_chart(
            worksheet=ws,
            data_range='A1:D7',
            chart_type='xlLineMarkers',
            position='F20',
            width=400,
            height=300
        )
        
        chart_mgr.set_chart_title(line_chart, '月度销售趋势')
        chart_mgr.set_axis_title(line_chart, 'x', '月份')
        chart_mgr.set_axis_title(line_chart, 'y', '销售额')
        chart_mgr.set_legend(line_chart, 'right')
        
        # 准备饼图数据
        pie_data = [
            ['产品', '总销售额'],
            ['产品A', 9100],
            ['产品B', 6300],
            ['产品C', 5200]
        ]
        
        basic.set_range_values(ws, 'A10:B13', pie_data)
        
        # 创建饼图
        pie_chart = chart_mgr.create_chart(
            worksheet=ws,
            data_range='A10:B13',
            chart_type='xlPie',
            position='N2',
            width=350,
            height=300
        )
        
        chart_mgr.set_chart_title(pie_chart, '产品销售占比')
        chart_mgr.set_data_labels(pie_chart, show_percentage=True, show_category=True)
        
        # 设置图表样式
        chart_mgr.set_chart_style(column_chart, 10)
        chart_mgr.set_chart_style(line_chart, 15)
        chart_mgr.set_chart_style(pie_chart, 8)
        
        # 导出图表
        chart_export_dir = OUTPUT_DIR / "charts"
        chart_export_dir.mkdir(exist_ok=True)
        
        chart_mgr.export_chart(column_chart, str(chart_export_dir / "column_chart.png"))
        chart_mgr.export_chart(line_chart, str(chart_export_dir / "line_chart.png"))
        chart_mgr.export_chart(pie_chart, str(chart_export_dir / "pie_chart.png"))
        
        # 获取图表列表
        charts = chart_mgr.get_chart_list(wb)
        logger.info(f"创建的图表: {len(charts)} 个")
        
        # 保存文件
        output_file = OUTPUT_DIR / "chart_operations_demo.xlsx"
        basic.save_workbook(wb, str(output_file))
        
        logger.info(f"图表演示完成，文件保存至: {output_file}")
        
    except Exception as e:
        logger.error(f"图表演示失败: {e}")
        raise


def demo_format_operations(excel_mgr: ExcelManager) -> None:
    """
    演示格式设置操作
    
    Args:
        excel_mgr: Excel 管理器
    """
    logger.info("=== 开始演示格式设置操作 ===")
    
    try:
        basic = ExcelBasic(excel_mgr)
        format_mgr = ExcelFormat(excel_mgr)
        
        # 创建工作簿
        wb = basic.create_workbook()
        ws = wb.ActiveSheet
        ws.Name = "格式演示"
        
        # 准备财务数据
        financial_data = [
            ['项目', '预算', '实际支出', '差异', '完成率', '状态'],
            ['市场推广', 50000, 45000, 5000, 0.9, '正常'],
            ['研发投入', 200000, 220000, -20000, 1.1, '超支'],
            ['人员成本', 300000, 285000, 15000, 0.95, '正常'],
            ['设备采购', 80000, 75000, 5000, 0.9375, '正常'],
            ['办公费用', 30000, 35000, -5000, 1.167, '超支']
        ]
        
        basic.set_range_values(ws, 'A1:F6', financial_data)
        
        # 应用标题样式
        format_mgr.apply_header_style(ws, 'A1:F1')
        
        # 设置数字格式
        format_mgr.apply_predefined_format(ws, 'B2:D6', 'currency')  # 金额列
        format_mgr.apply_predefined_format(ws, 'E2:E6', 'percentage')  # 完成率
        
        # 设置条件格式
        # 差异列：负数显示红色
        format_mgr.create_conditional_format(
            worksheet=ws,
            range_address='D2:D6',
            condition_type='cell_value',
            condition_value=('less', 0),
            format_style={'font_color': 'RED', 'bold': True}
        )
        
        # 完成率列：色阶显示
        format_mgr.create_conditional_format(
            worksheet=ws,
            range_address='E2:E6',
            condition_type='color_scale',
            format_style={'colors': ['RED', 'YELLOW', 'GREEN']}
        )
        
        # 状态列：条件格式
        format_mgr.create_conditional_format(
            worksheet=ws,
            range_address='F2:F6',
            condition_type='cell_value',
            condition_value=('equal', '超支'),
            format_style={'font_color': 'WHITE', 'fill_color': 'RED', 'bold': True}
        )
        
        format_mgr.create_conditional_format(
            worksheet=ws,
            range_address='F2:F6',
            condition_type='cell_value',
            condition_value=('equal', '正常'),
            format_style={'font_color': 'WHITE', 'fill_color': 'GREEN'}
        )
        
        # 创建表格样式
        format_mgr.create_table_style(ws, 'A1:F6', 'TableStyleMedium15')
        
        # 添加汇总行
        summary_data = [
            ['总计', '=SUM(B2:B6)', '=SUM(C2:C6)', '=SUM(D2:D6)', '=AVERAGE(E2:E6)', '']
        ]
        basic.set_range_values(ws, 'A7:F7', summary_data)
        
        # 设置汇总行格式
        format_mgr.set_font(ws, 'A7:F7', bold=True, font_size=12)
        format_mgr.set_fill(ws, 'A7:F7', fill_color='LIGHT_GRAY')
        format_mgr.set_borders(ws, 'A7:F7', border_weight='xlMedium')
        
        # 自动调整列宽
        basic.auto_fit_columns(ws)
        
        # 保存文件
        output_file = OUTPUT_DIR / "format_operations_demo.xlsx"
        basic.save_workbook(wb, str(output_file))
        
        logger.info(f"格式设置演示完成，文件保存至: {output_file}")
        
    except Exception as e:
        logger.error(f"格式设置演示失败: {e}")
        raise


def demo_print_operations(excel_mgr: ExcelManager) -> None:
    """
    演示打印操作
    
    Args:
        excel_mgr: Excel 管理器
    """
    logger.info("=== 开始演示打印操作 ===")
    
    try:
        basic = ExcelBasic(excel_mgr)
        print_mgr = ExcelPrint(excel_mgr)
        format_mgr = ExcelFormat(excel_mgr)
        
        # 创建工作簿
        wb = basic.create_workbook()
        ws = wb.ActiveSheet
        ws.Name = "打印演示"
        
        # 准备报表数据
        report_data = [
            ['月度销售报表'],
            [''],
            ['产品类别', '1月', '2月', '3月', '总计'],
            ['电子产品', 15000, 18000, 16000, 49000],
            ['服装鞋帽', 12000, 14000, 13000, 39000],
            ['家居用品', 8000, 9000, 10000, 27000],
            ['食品饮料', 6000, 7000, 8000, 21000],
            ['', '', '', '', ''],
            ['总计', 41000, 48000, 47000, 136000]
        ]
        
        basic.set_range_values(ws, 'A1:E9', report_data)
        
        # 设置报表格式
        # 标题
        format_mgr.set_font(ws, 'A1', font_size=16, bold=True)
        format_mgr.set_alignment(ws, 'A1', horizontal='center')
        ws.Range('A1:E1').Merge()
        
        # 表头
        format_mgr.apply_header_style(ws, 'A3:E3')
        
        # 数据格式
        format_mgr.apply_predefined_format(ws, 'B4:E9', 'currency')
        
        # 总计行
        format_mgr.set_font(ws, 'A9:E9', bold=True)
        format_mgr.set_fill(ws, 'A9:E9', fill_color='LIGHT_GRAY')
        
        # 设置边框
        format_mgr.set_borders(ws, 'A3:E9')
        
        # 应用打印模板
        print_mgr.apply_print_template(ws, 'default')
        
        # 设置打印区域
        print_mgr.set_print_area(ws, 'A1:E9')
        
        # 设置页眉页脚
        print_mgr.set_headers_footers(
            worksheet=ws,
            center_header='公司月度销售报表',
            left_footer='机密文件',
            center_footer='第 &P 页，共 &N 页',
            right_footer='&D &T'
        )
        
        # 设置打印选项
        print_mgr.set_print_options(
            worksheet=ws,
            print_gridlines=True,
            print_headings=False
        )
        
        # 导出为PDF
        pdf_file = OUTPUT_DIR / "sales_report.pdf"
        print_mgr.export_to_pdf(ws, str(pdf_file), quality='standard')
        
        # 保存文件
        output_file = OUTPUT_DIR / "print_operations_demo.xlsx"
        basic.save_workbook(wb, str(output_file))
        
        logger.info(f"打印演示完成，文件保存至: {output_file}")
        logger.info(f"PDF报表导出至: {pdf_file}")
        
    except Exception as e:
        logger.error(f"打印演示失败: {e}")
        raise


def demo_comprehensive_example(excel_mgr: ExcelManager) -> None:
    """
    综合示例：创建完整的销售分析报告
    
    Args:
        excel_mgr: Excel 管理器
    """
    logger.info("=== 开始综合示例：销售分析报告 ===")
    
    try:
        basic = ExcelBasic(excel_mgr)
        format_mgr = ExcelFormat(excel_mgr)
        chart_mgr = ExcelChart(excel_mgr)
        pivot = ExcelPivot(excel_mgr)
        print_mgr = ExcelPrint(excel_mgr)
        
        # 创建工作簿
        wb = basic.create_workbook()
        
        # 1. 创建原始数据工作表
        data_ws = wb.ActiveSheet
        data_ws.Name = "原始数据"
        
        # 生成模拟销售数据
        import random
        from datetime import datetime, timedelta
        
        products = ['笔记本电脑', '台式机', '显示器', '键盘', '鼠标', '音响']
        regions = ['北京', '上海', '广州', '深圳', '杭州', '南京']
        salespeople = ['张三', '李四', '王五', '赵六', '钱七', '孙八']
        
        sales_data = [['日期', '产品', '地区', '销售员', '数量', '单价', '总额']]
        
        start_date = datetime(2023, 1, 1)
        for i in range(100):  # 生成100条记录
            date = start_date + timedelta(days=random.randint(0, 365))
            product = random.choice(products)
            region = random.choice(regions)
            salesperson = random.choice(salespeople)
            quantity = random.randint(1, 20)
            
            # 根据产品设置价格范围
            price_ranges = {
                '笔记本电脑': (3000, 8000),
                '台式机': (2000, 6000),
                '显示器': (800, 3000),
                '键盘': (50, 500),
                '鼠标': (30, 300),
                '音响': (100, 1000)
            }
            
            min_price, max_price = price_ranges[product]
            unit_price = random.randint(min_price, max_price)
            total = quantity * unit_price
            
            sales_data.append([
                date.strftime('%Y-%m-%d'),
                product, region, salesperson,
                quantity, unit_price, total
            ])
        
        basic.set_range_values(data_ws, 'A1:G101', sales_data)
        
        # 格式化原始数据
        format_mgr.apply_header_style(data_ws, 'A1:G1')
        format_mgr.apply_predefined_format(data_ws, 'E2:G101', 'integer')
        format_mgr.apply_predefined_format(data_ws, 'F2:F101', 'currency')
        format_mgr.apply_predefined_format(data_ws, 'G2:G101', 'currency')
        basic.auto_fit_columns(data_ws)
        
        # 2. 创建数据透视表工作表
        pivot_ws = basic.add_worksheet(wb, "数据透视表")
        
        # 按产品和地区的销售透视表
        pivot_table1 = pivot.create_pivot_table(
            source_worksheet=data_ws,
            source_range='A1:G101',
            target_worksheet=pivot_ws,
            target_cell='A1',
            table_name='产品地区销售透视表'
        )
        
        pivot.add_row_field(pivot_table1, '产品')
        pivot.add_column_field(pivot_table1, '地区')
        pivot.add_data_field(pivot_table1, '总额', 'xlSum', '销售总额')
        pivot.format_pivot_table(pivot_table1, 'TableStyleMedium9')
        
        # 按销售员的业绩透视表
        pivot_table2 = pivot.create_pivot_table(
            source_worksheet=data_ws,
            source_range='A1:G101',
            target_worksheet=pivot_ws,
            target_cell='A15',
            table_name='销售员业绩透视表'
        )
        
        pivot.add_row_field(pivot_table2, '销售员')
        pivot.add_data_field(pivot_table2, '总额', 'xlSum', '销售总额')
        pivot.add_data_field(pivot_table2, '数量', 'xlSum', '销售数量')
        pivot.format_pivot_table(pivot_table2, 'TableStyleMedium6')
        
        # 3. 创建图表工作表
        chart_ws = basic.add_worksheet(wb, "图表分析")
        
        # 准备图表数据（从透视表获取）
        pivot_data = pivot.get_pivot_table_data(pivot_table1)
        
        # 简化数据用于图表
        chart_data = [
            ['产品', '总销售额'],
            ['笔记本电脑', 0],
            ['台式机', 0],
            ['显示器', 0],
            ['键盘', 0],
            ['鼠标', 0],
            ['音响', 0]
        ]
        
        # 计算各产品总销售额
        product_totals = {}
        for row in sales_data[1:]:
            product = row[1]
            total = row[6]
            if product in product_totals:
                product_totals[product] += total
            else:
                product_totals[product] = total
        
        for i, product in enumerate(['笔记本电脑', '台式机', '显示器', '键盘', '鼠标', '音响']):
            chart_data[i + 1][1] = product_totals.get(product, 0)
        
        basic.set_range_values(chart_ws, 'A1:B7', chart_data)
        
        # 创建柱状图
        column_chart = chart_mgr.create_chart(
            worksheet=chart_ws,
            data_range='A1:B7',
            chart_type='xlColumnClustered',
            position='D2',
            width=500,
            height=350
        )
        
        chart_mgr.set_chart_title(column_chart, '各产品销售额对比')
        chart_mgr.set_axis_title(column_chart, 'x', '产品类别')
        chart_mgr.set_axis_title(column_chart, 'y', '销售额（元）')
        chart_mgr.set_chart_style(column_chart, 10)
        
        # 创建饼图
        pie_chart = chart_mgr.create_chart(
            worksheet=chart_ws,
            data_range='A1:B7',
            chart_type='xlPie',
            position='D20',
            width=400,
            height=350
        )
        
        chart_mgr.set_chart_title(pie_chart, '产品销售占比')
        chart_mgr.set_data_labels(pie_chart, show_percentage=True, show_category=True)
        chart_mgr.set_chart_style(pie_chart, 8)
        
        # 4. 创建报告工作表
        report_ws = basic.add_worksheet(wb, "销售报告")
        
        # 报告标题
        report_title = [['2023年度销售分析报告']]
        basic.set_range_values(report_ws, 'A1:A1', report_title)
        
        # 设置标题格式
        format_mgr.set_font(report_ws, 'A1', font_size=18, bold=True)
        format_mgr.set_alignment(report_ws, 'A1', horizontal='center')
        report_ws.Range('A1:F1').Merge()
        
        # 添加报告内容
        report_content = [
            [''],
            ['报告摘要：'],
            ['1. 总销售记录：100条'],
            ['2. 涉及产品：6类'],
            ['3. 覆盖地区：6个城市'],
            ['4. 销售团队：6人'],
            [''],
            ['主要发现：'],
            ['• 笔记本电脑是主要销售产品'],
            ['• 北京和上海是主要销售市场'],
            ['• 销售业绩在各销售员之间分布相对均匀'],
            [''],
            ['建议：'],
            ['• 加大笔记本电脑的库存投入'],
            ['• 在北京和上海增加销售人员'],
            ['• 对表现优秀的销售员给予奖励']
        ]
        
        for i, content in enumerate(report_content):
            basic.set_range_values(report_ws, f'A{i+2}:A{i+2}', [content])
        
        # 设置报告格式
        format_mgr.set_font(report_ws, 'A3', bold=True, font_size=14)
        format_mgr.set_font(report_ws, 'A8', bold=True, font_size=14)
        format_mgr.set_font(report_ws, 'A12', bold=True, font_size=14)
        
        # 5. 设置打印格式
        for ws in [data_ws, pivot_ws, chart_ws, report_ws]:
            print_mgr.apply_print_template(ws, 'default')
            print_mgr.set_headers_footers(
                worksheet=ws,
                center_header='销售分析报告',
                right_footer='&D &T'
            )
        
        # 6. 保存文件
        output_file = OUTPUT_DIR / "comprehensive_sales_analysis.xlsx"
        basic.save_workbook(wb, str(output_file))
        
        # 7. 导出报告为PDF
        pdf_file = OUTPUT_DIR / "sales_analysis_report.pdf"
        print_mgr.export_workbook_to_pdf(wb, str(pdf_file))
        
        logger.info(f"综合销售分析报告完成！")
        logger.info(f"Excel文件: {output_file}")
        logger.info(f"PDF报告: {pdf_file}")
        
    except Exception as e:
        logger.error(f"综合示例失败: {e}")
        raise


def main():
    """
    主函数
    """
    logger.info("=== Excel 自动化操作演示开始 ===")
    
    try:
        # 确保输出目录存在
        OUTPUT_DIR.mkdir(exist_ok=True)
        
        # 使用 Excel 管理器
        with ExcelManager(visible=True, alerts=False) as excel_mgr:
            logger.info("Excel 应用程序初始化成功")
            
            # 演示各个功能模块
            demo_basic_operations(excel_mgr)
            time.sleep(1)
            
            demo_format_operations(excel_mgr)
            time.sleep(1)
            
            demo_chart_operations(excel_mgr)
            time.sleep(1)
            
            demo_pivot_operations(excel_mgr)
            time.sleep(1)
            
            demo_print_operations(excel_mgr)
            time.sleep(1)
            
            # 尝试演示宏操作（可能需要特殊权限）
            try:
                demo_macro_operations(excel_mgr)
            except Exception as e:
                logger.warning(f"宏操作演示跳过（可能需要启用宏支持）: {e}")
            
            time.sleep(1)
            
            # 综合示例
            demo_comprehensive_example(excel_mgr)
            
        logger.info("=== 所有演示完成 ===")
        logger.info(f"输出文件保存在: {OUTPUT_DIR}")
        
        # 显示生成的文件列表
        output_files = list(OUTPUT_DIR.glob("*"))
        logger.info("生成的文件：")
        for file in output_files:
            logger.info(f"  - {file.name}")
        
    except Exception as e:
        logger.error(f"程序执行失败: {e}")
        raise


if __name__ == "__main__":
    try:
        main()
        print("\n程序执行完成！请查看输出目录中的文件。")
        input("按回车键退出...")
    except KeyboardInterrupt:
        print("\n程序被用户中断")
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        input("按回车键退出...")