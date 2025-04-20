#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
from setuptools import setup, find_packages

# 读取版本信息
with open(os.path.join('kpichecker', '__init__.py'), 'r', encoding='utf-8') as f:
    for line in f:
        if line.startswith('__version__'):
            version = line.split('=')[1].strip().strip("'").strip('"')
            break
    else:
        version = '0.1.0'

# 读取README作为长描述
with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='kpichecker',
    version=version,
    description='检查Word文档中的KPI自评语句是否存在',
    long_description=long_description,
    long_description_content_type='text/markdown',
    author='KPI Checker Team',
    author_email='example@example.com',
    url='https://github.com/yourusername/kpichecker',
    packages=find_packages(),
    install_requires=[
        'python-docx>=0.8.11',
        'pandas>=1.3.0',
        'openpyxl>=3.0.7',
    ],
    entry_points={
        'console_scripts': [
            'kpichecker=kpichecker.main:main',
        ],
    },
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: End Users/Desktop',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Operating System :: OS Independent',
        'Topic :: Office/Business',
        'Topic :: Utilities',
    ],
    python_requires='>=3.6',
    keywords='kpi, word, docx, checker',
    project_urls={
        'Bug Reports': 'https://github.com/yourusername/kpichecker/issues',
        'Source': 'https://github.com/yourusername/kpichecker',
    },
)
