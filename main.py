#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
KPI检查器主程序
"""

import os
import sys
import argparse
import json
import datetime
import warnings
from typing import Dict, List, Any, Optional

# 检查依赖
def check_dependencies():
    """检查必要的依赖是否已安装"""
    missing_deps = []
    
    try:
        import docx
    except ImportError:
        missing_deps.append("python-docx")
        
    try:
        import pandas
    except ImportError:
        missing_deps.append("pandas")
        
    try:
        import openpyxl
    except ImportError:
        missing_deps.append("openpyxl")
        
    if missing_deps:
        print("警告: 以下依赖库未安装，某些功能可能不可用:")
        for dep in missing_deps:
            print(f"  - {dep}")
        print("\n您可以通过以下命令安装所有依赖:")
        print("pip install -r requirements.txt")
        print("或者单独安装缺失的库:")
        print(f"pip install {' '.join(missing_deps)}")
        print("\n程序将尝试继续运行，但可能会出现错误。\n")
        
    return len(missing_deps) == 0

# 主要逻辑导入放在依赖检查之后
from kpichecker.config import Config
from kpichecker.docx_scanner import DocxScanner
from kpichecker.docx_parser import DocxParser
from kpichecker.kpi_checker import KPIChecker
from kpichecker.report_generator import ReportGenerator
from kpichecker.file_fixer import FileFixer


def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description='KPI自评检查工具')
    
    # 基本操作
    parser.add_argument('--dir', '-d', type=str, default='.',
                        help='要扫描的目录路径，默认为当前目录')
    parser.add_argument('--recursive', '-r', action='store_true',
                        help='是否递归扫描子目录')
    parser.add_argument('--output', '-o', type=str, default='.',
                        help='报告输出目录，默认为当前目录')
                        
    # 报告格式
    parser.add_argument('--format', '-f', type=str, choices=['console', 'txt', 'excel'],
                        help='报告格式（console, txt, excel），默认使用配置文件中的设置')
                        
    # 配置选项
    parser.add_argument('--config', '-c', type=str,
                        help='配置文件路径')
    parser.add_argument('--paragraphs', '-p', type=int,
                        help='要检查的最后几个段落数量')
                        
    # 修复选项
    parser.add_argument('--fix', action='store_true',
                        help='修复缺少KPI自评的文件')
    parser.add_argument('--fix-list', type=str,
                        help='要修复的文件列表（JSON文件路径）')
    parser.add_argument('--no-backup', action='store_true',
                        help='修复时不创建备份文件')
    parser.add_argument('--fix-preview', action='store_true',
                        help='仅预览修复效果，不实际修改文件')
                        
    # 其他选项
    parser.add_argument('--version', '-v', action='store_true',
                        help='显示版本信息')
    parser.add_argument('--check-deps', action='store_true',
                        help='检查依赖库是否已安装')
                        
    return parser.parse_args()


def main():
    """主函数"""
    args = parse_arguments()
    
    # 检查依赖
    if args.check_deps:
        all_deps_installed = check_dependencies()
        return 0 if all_deps_installed else 1
    else:
        # 静默检查依赖
        check_dependencies()
    
    # 显示版本信息
    if args.version:
        from kpichecker import __version__, __description__
        print(f"KPI检查器 v{__version__}")
        print(__description__)
        return 0
        
    # 加载配置
    config = Config(args.config)
    
    # 更新配置
    if args.paragraphs:
        config.set('check_last_paragraphs', args.paragraphs)
        
    # 创建工具实例
    scanner = DocxScanner(recursive=args.recursive)
    parser = DocxParser(last_paragraphs=config.get('check_last_paragraphs', 10))
    checker = KPIChecker(config=config)
    report_gen = ReportGenerator(config=config)
    fixer = FileFixer(config=config)
    
    # 如果只是修复指定的文件列表
    if args.fix_list:
        try:
            with open(args.fix_list, 'r', encoding='utf-8') as f:
                files_to_fix = json.load(f)
                
            if not isinstance(files_to_fix, list):
                print(f"错误: 文件列表格式不正确，应为JSON数组")
                return 1
                
            # 修复文件
            fix_results = fixer.fix_multiple_files(files_to_fix, not args.no_backup)
            
            # 输出结果
            success_count = sum(1 for r in fix_results.values() if r.get('success', False))
            print(f"修复完成。成功: {success_count}/{len(files_to_fix)}")
            
            return 0
            
        except Exception as e:
            print(f"修复文件时出错: {str(e)}")
            return 1
            
    # 扫描文件
    print(f"正在扫描目录 {args.dir}...")
    docx_files = scanner.scan_directory(args.dir)
    
    if not docx_files:
        print(f"没有找到.docx文件")
        return 0
        
    print(f"找到 {len(docx_files)} 个.docx文件")
    
    # 解析文件并检查KPI自评
    print("正在分析文件...")
    
    file_contents = {}
    file_data = {}
    
    for file_path in docx_files:
        print(f"分析文件: {os.path.basename(file_path)}")
        
        # 提取KPI相关文本
        kpi_text, doc_info = parser.extract_kpi_section(file_path)
        
        if 'error' in doc_info:
            print(f"  错误: {doc_info['error']}")
            continue
            
        file_contents[file_path] = kpi_text
        file_data[file_path] = doc_info
        
    # 检查KPI自评语句
    print("检查KPI自评语句...")
    results = checker.check_files(file_contents)
    
    # 分析结果
    analysis = checker.analyze_results(results)
    
    print(f"总文件数: {analysis['total_files']}")
    print(f"包含KPI自评的文件: {analysis['files_with_kpi']}")
    print(f"缺少KPI自评的文件: {analysis['files_without_kpi']}")
    
    if analysis['files_with_kpi'] > 0:
        print(f"平均KPI分数: {analysis['average_score']}")
    
    try:    
        # 生成报告
        print("生成报告...")
        report_path = report_gen.generate_report(
            results,
            file_data,
            args.output,
            args.format
        )
        
        if args.format != 'console':
            print(f"报告已保存到: {report_path}")
    except Exception as e:
        print(f"生成报告时出错: {str(e)}")
        print("尝试降级为文本报告...")
        try:
            report_path = report_gen.generate_report(
                results,
                file_data,
                args.output,
                'txt'  # 强制使用文本格式
            )
            print(f"文本报告已保存到: {report_path}")
        except Exception as e2:
            print(f"无法生成任何报告: {str(e2)}")
            return 1
        
    # 修复缺少KPI自评的文件
    if args.fix or args.fix_preview:
        files_to_fix = [
            file_path for file_path, result in results.items()
            if not result.get('has_kpi_statement', False)
        ]
        
        if not files_to_fix:
            print("没有需要修复的文件")
        else:
            print(f"找到 {len(files_to_fix)} 个需要修复的文件")
            
            if args.fix_preview:
                # 只预览修复效果
                print("\n修复预览:")
                print("-" * 80)
                
                for file_path in files_to_fix:
                    preview = fixer.generate_fix_preview(file_path)
                    if preview.get('success', False):
                        print(f"文件: {preview['file_name']}")
                        print(f"将添加: {preview['template_to_add']}")
                        print(f"位置: {preview['position']}")
                        print("")
                        
            elif args.fix:
                # 实际修复文件
                print(f"\n正在修复 {len(files_to_fix)} 个文件...")
                fix_results = fixer.fix_multiple_files(files_to_fix, not args.no_backup)
                
                success_count = sum(1 for r in fix_results.values() if r.get('success', False))
                print(f"修复完成。成功: {success_count}/{len(files_to_fix)}")
                
                # 保存修复的文件列表，以便之后回滚
                fix_list_path = os.path.join(args.output, 'fixed_files.json')
                with open(fix_list_path, 'w', encoding='utf-8') as f:
                    json.dump(files_to_fix, f, ensure_ascii=False, indent=4)
                    
                print(f"已保存修复文件列表到: {fix_list_path}")
                
    return 0


if __name__ == '__main__':
    sys.exit(main())
