#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
报告生成模块，用于生成检查结果报告
"""

import os
import datetime
import sys
import warnings
from typing import Dict, List, Any, Optional

try:
    import pandas as pd
    pandas_available = True
except ImportError:
    pandas_available = False
    warnings.warn("pandas库未安装，Excel报告功能将不可用")

try:
    import openpyxl
    openpyxl_available = True
except ImportError:
    openpyxl_available = False
    warnings.warn("openpyxl库未安装，Excel报告功能将不可用")

from .config import Config


class ReportGenerator:
    """报告生成器，用于生成检查结果报告"""
    
    def __init__(self, config: Config = None):
        """
        初始化报告生成器
        
        Args:
            config: 配置对象，如果为None则使用默认配置
        """
        self.config = config or Config()
        self.report_format = self.config.get('report', {}).get('default_format', 'excel')
        
    def generate_report(self, 
                         results: Dict[str, Dict[str, Any]], 
                         file_data: Dict[str, Dict[str, Any]], 
                         output_directory: str = '.', 
                         report_format: Optional[str] = None) -> str:
        """
        生成检查结果报告
        
        Args:
            results: 检查结果字典
            file_data: 文件数据字典
            output_directory: 输出目录
            report_format: 报告格式，可选值为 'console', 'txt', 'excel'
                如果为None，则使用配置中的默认格式
                
        Returns:
            报告文件路径或控制台输出的文本
        """
        # 使用指定的格式或默认格式
        format_to_use = report_format or self.report_format
        
        # 如果指定Excel格式但缺少依赖，自动降级为文本格式
        if format_to_use == 'excel' and (not pandas_available or not openpyxl_available):
            print("警告: Excel报告需要pandas和openpyxl库，但未安装。自动降级为文本报告格式。")
            format_to_use = 'txt'
            
        if format_to_use == 'console':
            return self._generate_console_report(results, file_data)
        elif format_to_use == 'txt':
            return self._generate_txt_report(results, file_data, output_directory)
        elif format_to_use == 'excel':
            # 此时已确保pandas和openpyxl可用
            return self._generate_excel_report(results, file_data, output_directory)
        else:
            raise ValueError(f"不支持的报告格式: {format_to_use}")
            
    def _generate_console_report(self, 
                                results: Dict[str, Dict[str, Any]], 
                                file_data: Dict[str, Dict[str, Any]]) -> str:
        """
        生成控制台报告
        
        Args:
            results: 检查结果字典
            file_data: 文件数据字典
            
        Returns:
            报告文本
        """
        report = []
        report.append("=" * 80)
        report.append("KPI自评检查报告")
        report.append("=" * 80)
        report.append(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append(f"总文件数: {len(results)}")
        
        # 统计包含KPI自评的文件数量
        files_with_kpi = sum(1 for r in results.values() if r.get('has_kpi_statement', False))
        files_without_kpi = len(results) - files_with_kpi
        
        report.append(f"包含KPI自评的文件数: {files_with_kpi}")
        report.append(f"缺少KPI自评的文件数: {files_without_kpi}")
        report.append("-" * 80)
        
        # 缺少KPI自评的文件
        if files_without_kpi > 0:
            report.append("\n缺少KPI自评的文件:")
            report.append("-" * 80)
            
            for file_path, result in results.items():
                if not result.get('has_kpi_statement', False):
                    file_name = os.path.basename(file_path)
                    report.append(f"文件: {file_name}")
                    report.append(f"路径: {file_path}")
                    report.append("")
                    
        # 包含KPI自评的文件及其分数
        if files_with_kpi > 0:
            report.append("\n包含KPI自评的文件及其分数:")
            report.append("-" * 80)
            
            for file_path, result in results.items():
                if result.get('has_kpi_statement', False):
                    file_name = os.path.basename(file_path)
                    score = result.get('score', 'N/A')
                    match_text = result.get('match_text', 'N/A')
                    
                    report.append(f"文件: {file_name}")
                    report.append(f"分数: {score}")
                    report.append(f"匹配文本: {match_text}")
                    report.append("")
                    
        # 计算平均分
        if files_with_kpi > 0:
            total_score = sum(int(r.get('score', 0)) for r in results.values() if r.get('score'))
            avg_score = total_score / files_with_kpi
            report.append(f"\n平均分: {avg_score:.2f}")
            
        return "\n".join(report)
        
    def _generate_txt_report(self, 
                            results: Dict[str, Dict[str, Any]], 
                            file_data: Dict[str, Dict[str, Any]], 
                            output_directory: str) -> str:
        """
        生成TXT报告
        
        Args:
            results: 检查结果字典
            file_data: 文件数据字典
            output_directory: 输出目录
            
        Returns:
            报告文件路径
        """
        # 生成控制台报告文本
        report_text = self._generate_console_report(results, file_data)
        
        # 获取输出文件名
        txt_filename = self.config.get('report', {}).get('txt_filename', 'KPI检查报告.txt')
        output_path = os.path.join(output_directory, txt_filename)
        
        # 写入文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report_text)
            
        return output_path
        
    def _generate_excel_report(self, 
                              results: Dict[str, Dict[str, Any]], 
                              file_data: Dict[str, Dict[str, Any]], 
                              output_directory: str) -> str:
        """
        生成Excel报告
        
        Args:
            results: 检查结果字典
            file_data: 文件数据字典
            output_directory: 输出目录
            
        Returns:
            报告文件路径
        """
        # 此函数只有在pandas和openpyxl已安装的情况下才会被调用
        
        # 准备数据
        report_data = []
        
        for file_path, result in results.items():
            file_name = os.path.basename(file_path)
            has_kpi = result.get('has_kpi_statement', False)
            score = result.get('score', 'N/A')
            match_text = result.get('match_text', 'N/A')
            
            report_data.append({
                '文件名': file_name,
                '文件路径': file_path,
                '包含KPI自评': '是' if has_kpi else '否',
                'KPI自评分数': score if has_kpi else 'N/A',
                '匹配文本': match_text if has_kpi else 'N/A'
            })
            
        # 创建DataFrame
        df = pd.DataFrame(report_data)
        
        # 获取输出文件名
        excel_filename = self.config.get('report', {}).get('excel_filename', 'KPI检查报告.xlsx')
        output_path = os.path.join(output_directory, excel_filename)
        
        try:
            # 写入Excel文件
            writer = pd.ExcelWriter(output_path, engine='openpyxl')
            
            # 写入摘要表
            summary_data = {
                '项目': ['总文件数', '包含KPI自评的文件数', '缺少KPI自评的文件数', '平均分'],
                '值': [
                    len(results),
                    sum(1 for r in results.values() if r.get('has_kpi_statement', False)),
                    sum(1 for r in results.values() if not r.get('has_kpi_statement', False)),
                    sum(float(r.get('score', 0)) for r in results.values() if r.get('score')) / 
                        sum(1 for r in results.values() if r.get('has_kpi_statement', False))
                        if sum(1 for r in results.values() if r.get('has_kpi_statement', False)) > 0 else 0
                ]
            }
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='摘要', index=False)
            
            # 写入详细数据
            df.to_excel(writer, sheet_name='详细数据', index=False)
            
            # 写入缺少KPI自评的文件
            missing_kpi_data = [
                {'文件名': os.path.basename(file_path), '文件路径': file_path}
                for file_path, result in results.items()
                if not result.get('has_kpi_statement', False)
            ]
            
            if missing_kpi_data:
                missing_df = pd.DataFrame(missing_kpi_data)
                missing_df.to_excel(writer, sheet_name='缺少KPI自评的文件', index=False)
                
            # 保存文件
            writer.close()
            
            return output_path
            
        except Exception as e:
            print(f"生成Excel报告时出错: {str(e)}")
            print("自动降级为文本报告格式。")
            return self._generate_txt_report(results, file_data, output_directory)
