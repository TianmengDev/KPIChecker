#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文档扫描模块，用于查找目录中的.docx文件
"""

import os
from typing import List, Optional


class DocxScanner:
    """文档扫描器，用于查找.docx文件"""
    
    def __init__(self, recursive: bool = False):
        """
        初始化文档扫描器
        
        Args:
            recursive: 是否递归扫描子目录
        """
        self.recursive = recursive
        
    def scan_directory(self, directory: str = '.') -> List[str]:
        """
        扫描指定目录，返回所有.docx文件的路径
        
        Args:
            directory: 要扫描的目录路径，默认为当前目录
            
        Returns:
            文件路径列表
        """
        docx_files = []
        
        # 确保目录路径存在
        if not os.path.exists(directory):
            print(f"目录 {directory} 不存在")
            return docx_files
            
        # 如果递归扫描
        if self.recursive:
            for root, _, files in os.walk(directory):
                for file in files:
                    if file.lower().endswith('.docx') and not file.startswith('~$'):  # 排除临时文件
                        docx_files.append(os.path.join(root, file))
        else:
            # 非递归只扫描当前目录
            for file in os.listdir(directory):
                if file.lower().endswith('.docx') and not file.startswith('~$'):
                    docx_files.append(os.path.join(directory, file))
                    
        return docx_files
        
    def get_file_info(self, file_path: str) -> Optional[dict]:
        """
        获取文件信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件信息字典，包含文件名、路径、大小等
        """
        if not os.path.exists(file_path):
            return None
            
        file_info = {
            'name': os.path.basename(file_path),
            'path': file_path,
            'size': os.path.getsize(file_path),
            'last_modified': os.path.getmtime(file_path)
        }
        
        return file_info
        
    def scan_with_info(self, directory: str = '.') -> List[dict]:
        """
        扫描目录并返回带有文件信息的结果
        
        Args:
            directory: 要扫描的目录路径
            
        Returns:
            文件信息字典列表
        """
        docx_files = self.scan_directory(directory)
        result = []
        
        for file_path in docx_files:
            file_info = self.get_file_info(file_path)
            if file_info:
                result.append(file_info)
                
        return result
