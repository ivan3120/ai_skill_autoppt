# -*- coding: utf-8 -*-
"""
幻灯片配置加载器
加载 configs/slide_templates.yaml 并提供访问接口
"""

import os
import yaml
from typing import Dict, List, Optional, Any
from pathlib import Path


class SlideConfig:
    """幻灯片配置类"""

    _instance = None
    _config = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if self._config is None:
            self._load()

    def _load(self):
        """加载配置文件"""
        # 查找配置文件
        config_path = self._find_config_path()
        if config_path and os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                self._config = yaml.safe_load(f)
        else:
            # 使用默认配置
            self._config = self._get_default_config()

    def _find_config_path(self) -> str:
        """查找配置文件路径"""
        # 当前目录
        if os.path.exists('configs/slide_templates.yaml'):
            return 'configs/slide_templates.yaml'
        # 项目根目录
        root = Path(__file__).parent.parent
        config_path = root / 'configs' / 'slide_templates.yaml'
        if config_path.exists():
            return str(config_path)
        return None

    def _get_default_config(self) -> Dict:
        """获取默认配置"""
        return {
            'slide_types': {
                'cover': {
                    'title': '云核心网网络架构评估报告',
                    'subtitle': '运营商客户汇报材料',
                    'date': '2026年4月'
                },
                'toc': {'title': '目录'},
                'intro': {'title': '一、评估方案介绍'},
                'background': {'title': '二、业务背景与网络概况'},
                'overview': {'title': '三、整体评估概览'},
                'dimension_detail': {'title_pattern': '{idx}、{name}'},
                'summary': {'title': '十一、总结与优化建议'}
            },
            'dimensions': {
                'MT0': {'name': '网元高稳评估', 'short_name': '网元高稳', 'idx': 4},
                'MT1': {'name': '部署离散度评估', 'short_name': '部署离散度', 'idx': 5},
                'MT2': {'name': '组网架构高可用评估', 'short_name': '组网架构', 'idx': 6},
                'MT3': {'name': '业务路由高可用评估', 'short_name': '业务路由', 'idx': 7},
                'MT4': {'name': '网络容量评估', 'short_name': '网络容量', 'idx': 8},
                'MT5': {'name': '网元版本EOS评估', 'short_name': '网元版本', 'idx': 9},
                'MT6': {'name': '云平台版本EOS评估', 'short_name': '云平台版本', 'idx': 10}
            }
        }

    def get_slide_type(self, slide_type: str) -> Dict:
        """获取幻灯片类型配置"""
        return self._config.get('slide_types', {}).get(slide_type, {})

    def get_dimension(self, dim_code: str) -> Dict:
        """获取评估维度配置"""
        return self._config.get('dimensions', {}).get(dim_code, {})

    def get_all_dimensions(self) -> List[Dict]:
        """获取所有评估维度配置"""
        dims = self._config.get('dimensions', {})
        return [{'code': k, **v} for k, v in dims.items()]

    def get_toc_template(self) -> List[str]:
        """获取目录模板"""
        return self._config.get('toc_template', [])

    def build_toc_items(self, selected_dimensions: List[str]) -> List[str]:
        """构建目录项"""
        template = self.get_toc_template()
        items = []
        for item in template:
            # 替换维度占位符
            if '{mt' in item:
                for dim in selected_dimensions:
                    dim_info = self.get_dimension(dim)
                    placeholder = f'{{{dim_info.get("short_name", dim).lower()}_name}}'
                    if placeholder in item:
                        item = item.replace(placeholder, f"（{dim_info.get('name', dim)}）")
            items.append(item)
        # 过滤掉未使用的占位符
        return [i for i in items if '{' not in i]

    def get_dimension_title(self, dim_code: str, selected_dims: List[str]) -> str:
        """获取维度详情页标题"""
        dim = self.get_dimension(dim_code)
        if not dim:
            return dim_code

        # 计算序号
        try:
            idx = selected_dims.index(dim_code) + 4  # 从第4页开始
        except:
            idx = dim.get('idx', 0)

        name = dim.get('name', dim_code)
        return f"{idx}、{name}"

    def get_status_for_risk(self, risk: str) -> str:
        """根据风险等级获取状态"""
        status_map = self._config.get('status_mapping', {
            '高': '需改进',
            '中': '待优化',
            '低': '良好'
        })
        return status_map.get(risk, '良好')


# 全局配置实例
config = SlideConfig()
