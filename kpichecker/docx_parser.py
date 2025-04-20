#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文档解析模块，用于读取和解析.docx文件内容
"""

import os
from typing import List, Dict, Any, Optional, Tuple

from docx import Document


class DocxParser:
    """文档解析器，用于读取和解析.docx文件内容"""
    
    def __init__(self, last_paragraphs: int = 10):
        """
        初始化文档解析器
        
        Args:
            last_paragraphs: 要提取的最后几个段落的数量
        """
        self.last_paragraphs = last_paragraphs
        
    def parse_file(self, file_path: str) -> Dict[str, Any]:
        """
        解析文档文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            解析结果字典，包含文件信息和内容
        """
        if not os.path.exists(file_path):
            return {'error': f"文件 {file_path} 不存在"}
            
        try:
            doc = Document(file_path)
            
            # 获取所有段落文本
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            
            # 提取文档的基本信息
            result = {
                'file_path': file_path,
                'file_name': os.path.basename(file_path),
                'total_paragraphs': len(paragraphs),
                'content': paragraphs,
            }
            
            # 提取最后几个段落
            last_n = min(self.last_paragraphs, len(paragraphs))
            result['last_paragraphs'] = paragraphs[-last_n:] if paragraphs else []
            
            # 连接最后几个段落形成一个文本块，方便后续处理
            result['last_paragraphs_text'] = '\n'.join(result['last_paragraphs'])
            
            return result
            
        except Exception as e:
            return {'error': f"解析文件 {file_path} 失败: {str(e)}"}
            
    def get_text_blocks(self, file_path: str, block_size: int = 5) -> List[str]:
        """
        将文档分割成文本块
        
        Args:
            file_path: 文件路径
            block_size: 每个块包含的段落数
            
        Returns:
            文本块列表
        """
        result = self.parse_file(file_path)
        
        if 'error' in result:
            return []
            
        paragraphs = result.get('content', [])
        
        # 将段落分组成块
        blocks = []
        for i in range(0, len(paragraphs), block_size):
            block = '\n'.join(paragraphs[i:i+block_size])
            blocks.append(block)
            
        return blocks
        
    def extract_kpi_section(self, file_path: str) -> Tuple[Optional[str], Optional[Dict[str, Any]]]:
        """
        提取文档中的KPI评估部分
        
        Args:
            file_path: 文件路径
            
        Returns:
            提取的KPI文本和文档信息元组
        """
        result = self.parse_file(file_path)
        
        if 'error' in result:
            return None, result
            
        # 我们假设KPI评估部分在文档的最后几个段落中
        kpi_text = result.get('last_paragraphs_text', '')
        
        return kpi_text, result
