#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文档修复模块，用于修复缺少KPI自评语句的文档
"""

import os
import shutil
from typing import Dict, Any, List, Optional, Tuple

from docx import Document

from .config import Config


class FileFixer:
    """文档修复器，用于修复缺少KPI自评语句的文档"""
    
    def __init__(self, config: Config = None):
        """
        初始化文档修复器
        
        Args:
            config: 配置对象，如果为None则使用默认配置
        """
        self.config = config or Config()
        self.template = self.config.get('fixer', {}).get('template', '{year}年第{quarter}季度KPI考核自评__分。')
        
    def fix_file(self, file_path: str, create_backup: bool = True) -> Dict[str, Any]:
        """
        修复文档，在文档末尾添加KPI自评语句模板
        
        Args:
            file_path: 文件路径
            create_backup: 是否创建备份文件
            
        Returns:
            修复结果字典
        """
        if not os.path.exists(file_path):
            return {'success': False, 'error': f"文件 {file_path} 不存在"}
            
        try:
            # 如果需要创建备份
            if create_backup:
                backup_path = f"{file_path}.bak"
                shutil.copy2(file_path, backup_path)
                
            # 打开文档
            doc = Document(file_path)
            
            # 获取当前年份和季度
            year_quarter = self.config.get_current_year_quarter()
            
            # 格式化模板
            template_text = self.template.format(
                year=year_quarter['year'],
                quarter=year_quarter['quarter']
            )
            
            # 在文档末尾添加段落
            doc.add_paragraph(template_text)
            
            # 保存文档
            doc.save(file_path)
            
            return {
                'success': True,
                'file_path': file_path,
                'template_added': template_text,
                'backup_created': create_backup,
                'backup_path': backup_path if create_backup else None
            }
            
        except Exception as e:
            return {'success': False, 'error': str(e)}
            
    def fix_multiple_files(self, file_paths: List[str], create_backup: bool = True) -> Dict[str, Dict[str, Any]]:
        """
        修复多个文档
        
        Args:
            file_paths: 文件路径列表
            create_backup: 是否创建备份文件
            
        Returns:
            文件路径到修复结果的映射
        """
        results = {}
        
        for file_path in file_paths:
            results[file_path] = self.fix_file(file_path, create_backup)
            
        return results
        
    def generate_fix_preview(self, file_path: str) -> Dict[str, Any]:
        """
        生成修复预览，不实际修改文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            预览信息字典
        """
        if not os.path.exists(file_path):
            return {'success': False, 'error': f"文件 {file_path} 不存在"}
            
        # 获取当前年份和季度
        year_quarter = self.config.get_current_year_quarter()
        
        # 格式化模板
        template_text = self.template.format(
            year=year_quarter['year'],
            quarter=year_quarter['quarter']
        )
        
        return {
            'success': True,
            'file_path': file_path,
            'file_name': os.path.basename(file_path),
            'template_to_add': template_text,
            'position': '文档末尾'
        }
        
    def restore_from_backup(self, file_path: str) -> Dict[str, Any]:
        """
        从备份文件恢复文档
        
        Args:
            file_path: 原始文件路径
            
        Returns:
            恢复结果字典
        """
        backup_path = f"{file_path}.bak"
        
        if not os.path.exists(backup_path):
            return {'success': False, 'error': f"备份文件 {backup_path} 不存在"}
            
        try:
            shutil.copy2(backup_path, file_path)
            return {'success': True, 'file_path': file_path, 'backup_path': backup_path}
        except Exception as e:
            return {'success': False, 'error': str(e)}
