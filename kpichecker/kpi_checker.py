#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
KPI检查模块，用于检查文档中是否包含KPI自评语句
"""

import re
from typing import Dict, List, Optional, Tuple, Any

from .config import Config


class KPIChecker:
    """KPI检查器，用于检查文档是否包含KPI自评语句"""
    
    def __init__(self, config: Config = None):
        """
        初始化KPI检查器
        
        Args:
            config: 配置对象，如果为None则使用默认配置
        """
        self.config = config or Config()
        self.patterns = self.config.get_kpi_patterns()
        
    def check_text(self, text: str) -> Dict[str, Any]:
        """
        检查文本是否包含KPI自评语句
        
        Args:
            text: 要检查的文本
            
        Returns:
            检查结果字典
        """
        if not text or not isinstance(text, str):
            return {
                'has_kpi_statement': False,
                'score': None,
                'matching_pattern': None,
                'match_text': None,
                'year': None
            }
            
        # 编译所有正则表达式
        compiled_patterns = [re.compile(pattern) for pattern in self.patterns]
        
        # 检查每个模式
        for i, pattern in enumerate(compiled_patterns):
            matches = pattern.finditer(text)
            
            for match in matches:
                match_text = match.group(0)
                
                # 根据不同的模式提取分数和年份
                if i == 0:  # 第一种模式: 带年份的完整模式
                    year = match.group(1)
                    score = match.group(2)
                    return {
                        'has_kpi_statement': True,
                        'score': score,
                        'matching_pattern': self.patterns[i],
                        'match_text': match_text,
                        'year': year
                    }
                else:  # 其他模式: 提取分数
                    score = match.group(1)
                    return {
                        'has_kpi_statement': True,
                        'score': score,
                        'matching_pattern': self.patterns[i],
                        'match_text': match_text,
                        'year': None
                    }
                    
        # 未找到匹配
        return {
            'has_kpi_statement': False,
            'score': None,
            'matching_pattern': None,
            'match_text': None,
            'year': None
        }
        
    def check_files(self, file_contents: Dict[str, str]) -> Dict[str, Dict[str, Any]]:
        """
        检查多个文件内容
        
        Args:
            file_contents: 文件路径到文件内容的映射
            
        Returns:
            文件路径到检查结果的映射
        """
        results = {}
        
        for file_path, content in file_contents.items():
            results[file_path] = self.check_text(content)
            
        return results
        
    def analyze_results(self, results: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """
        分析检查结果
        
        Args:
            results: 文件路径到检查结果的映射
            
        Returns:
            分析结果字典
        """
        total_files = len(results)
        files_with_kpi = sum(1 for result in results.values() if result.get('has_kpi_statement', False))
        files_without_kpi = total_files - files_with_kpi
        
        # 分数分布统计
        score_distribution = {}
        for result in results.values():
            score = result.get('score')
            if score:
                score = int(score)
                if score in score_distribution:
                    score_distribution[score] += 1
                else:
                    score_distribution[score] = 1
                    
        # 计算平均分
        total_score = sum(int(result.get('score', 0)) for result in results.values() if result.get('score'))
        avg_score = total_score / files_with_kpi if files_with_kpi > 0 else 0
        
        return {
            'total_files': total_files,
            'files_with_kpi': files_with_kpi,
            'files_without_kpi': files_without_kpi,
            'score_distribution': score_distribution,
            'average_score': round(avg_score, 2)
        }
