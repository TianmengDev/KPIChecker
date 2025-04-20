#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
配置模块，包含KPI检查器的配置选项和正则表达式模式
"""

import json
import os
import re
from typing import Dict, List, Any


class Config:
    """配置类，管理KPI检查器的设置"""

    def __init__(self, config_file: str = None):
        """
        初始化配置对象
        
        Args:
            config_file: 配置文件路径，如果为None则使用默认配置
        """
        # 默认配置
        self.default_config = {
            # 要检查的文档末尾段落数量
            "check_last_paragraphs": 10,
            
            # KPI自评正则表达式模式列表
            "kpi_patterns": [
                r"(\d{4})年?第?[一二三四]季度KPI考核自评(\d{1,3})分",  # 2024年第四季度KPI考核自评85分
                r"季度KPI考核自评(\d{1,3})分",                        # 季度KPI考核自评96分
                r"KPI\s*自评得分\s*(\d{1,3})\s*分",                  # KPI 自评得分 94 分
                r"自评得分\s*(\d{1,3})\s*分",                        # 自评得分 94 分
                r"KPI考核自评(\d{1,3})分",                           # KPI考核自评82分
            ],
            
            # 报告设置
            "report": {
                "default_format": "excel",  # 可选: console, txt, excel
                "excel_filename": "KPI检查报告.xlsx",
                "txt_filename": "KPI检查报告.txt",
            },
            
            # 修复设置
            "fixer": {
                "enabled": False,
                "template": "{year}年第{quarter}季度KPI考核自评__分。",
            }
        }
        
        self.config = self.default_config.copy()
        
        # 如果提供了配置文件，则读取它
        if config_file and os.path.exists(config_file):
            self.load_config(config_file)
            
    def load_config(self, config_file: str) -> None:
        """
        从文件加载配置
        
        Args:
            config_file: 配置文件路径
        """
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
                
            # 更新配置，但保留默认值（如果用户配置中没有指定）
            for key, value in user_config.items():
                if isinstance(value, dict) and key in self.config and isinstance(self.config[key], dict):
                    self.config[key].update(value)
                else:
                    self.config[key] = value
                    
        except (json.JSONDecodeError, IOError) as e:
            print(f"加载配置文件失败: {str(e)}")
            
    def save_config(self, config_file: str) -> None:
        """
        保存配置到文件
        
        Args:
            config_file: 配置文件路径
        """
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
                
        except IOError as e:
            print(f"保存配置文件失败: {str(e)}")
            
    def get(self, key: str, default: Any = None) -> Any:
        """
        获取配置值
        
        Args:
            key: 配置键
            default: 如果键不存在，返回的默认值
            
        Returns:
            配置值
        """
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """
        设置配置值
        
        Args:
            key: 配置键
            value: 配置值
        """
        self.config[key] = value
        
    def get_kpi_patterns(self) -> List[str]:
        """
        获取KPI自评正则表达式模式列表
        
        Returns:
            正则表达式模式列表
        """
        return self.config.get("kpi_patterns", [])
    
    def add_kpi_pattern(self, pattern: str) -> None:
        """
        添加KPI自评正则表达式模式
        
        Args:
            pattern: 正则表达式模式
        """
        if "kpi_patterns" not in self.config:
            self.config["kpi_patterns"] = []
            
        if pattern not in self.config["kpi_patterns"]:
            self.config["kpi_patterns"].append(pattern)
            
    def get_current_year_quarter(self) -> Dict[str, str]:
        """
        获取当前年份和季度
        
        Returns:
            包含年份和季度的字典
        """
        import datetime
        
        now = datetime.datetime.now()
        year = str(now.year)
        month = now.month
        
        if 1 <= month <= 3:
            quarter = "一"
        elif 4 <= month <= 6:
            quarter = "二"
        elif 7 <= month <= 9:
            quarter = "三"
        else:
            quarter = "四"
            
        return {"year": year, "quarter": quarter}


# 创建默认配置实例
default_config = Config()
