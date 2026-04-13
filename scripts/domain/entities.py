# -*- coding: utf-8 -*-
"""
AutoPPT 领域层 - 实体定义
核心业务对象
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, List, Dict


@dataclass
class EvaluationDimension:
    """评估维度实体"""
    code: str
    name: str
    description: str = ""

    def to_dict(self):
        return {"code": self.code, "name": self.name, "description": self.description}


@dataclass
class ProductDomain:
    """产品域实体"""
    name: str
    ne_count: int = 0
    dr_count: int = 0

    def to_dict(self):
        return {"name": self.name, "ne_count": self.ne_count, "dr_count": self.dr_count}


@dataclass
class NetworkElement:
    """网元实体"""
    name: str
    product_domain: str
    disaster_recovery_group: str
    status: str = "normal"

    def to_dict(self):
        return {
            "name": self.name,
            "product_domain": self.product_domain,
            "disaster_recovery_group": self.disaster_recovery_group,
            "status": self.status
        }


@dataclass
class EvaluationItem:
    """评估项实体"""
    code: str
    name: str
    dimension: str
    product_domain: str
    result: str = ""  # 通过/待改进/不通过
    suggestions: str = ""

    def to_dict(self):
        return {
            "code": self.code,
            "name": self.name,
            "dimension": self.dimension,
            "product_domain": self.product_domain,
            "result": self.result,
            "suggestions": self.suggestions
        }


@dataclass
class OntologyModel:
    """本体模型 - 领域对象"""
    product_domains: List[ProductDomain] = field(default_factory=list)
    data_centers: List[str] = field(default_factory=list)
    network_elements: List[NetworkElement] = field(default_factory=list)
    evaluation_items: List[EvaluationItem] = field(default_factory=list)
    dimensions: List[EvaluationDimension] = field(default_factory=list)
    overview_results: List[Dict] = field(default_factory=list)  # 整体评估概览数据
    parse_time: str = ""
    source_file: str = ""

    def to_dict(self):
        return {
            "产品域列表": [pd.to_dict() for pd in self.product_domains],
            "数据中心列表": self.data_centers,
            "网元列表": [ne.to_dict() for ne in self.network_elements],
            "评估项列表": [ei.to_dict() for ei in self.evaluation_items],
            "评估维度": [d.to_dict() for d in self.dimensions],
            "整体评估概览": self.overview_results,
            "解析时间": self.parse_time,
            "文件路径": self.source_file
        }

    @property
    def domain_count(self) -> int:
        return len(self.product_domains)

    @property
    def ne_count(self) -> int:
        return len(self.network_elements)

    @property
    def item_count(self) -> int:
        return len(self.evaluation_items)


@dataclass
class PPTContent:
    """PPT内容 - 领域对象"""
    title: str
    subtitle: str = ""
    dimensions: List[str] = field(default_factory=list)
    domains: List[str] = field(default_factory=list)
    slides: List[dict] = field(default_factory=list)
    style_config: dict = field(default_factory=dict)

    def to_dict(self):
        return {
            "title": self.title,
            "subtitle": self.subtitle,
            "dimensions": self.dimensions,
            "domains": self.domains,
            "slides": self.slides,
            "style_config": self.style_config
        }


@dataclass
class ReviewResult:
    """审核结果 - 值对象"""
    status: str  # passed/needs_revision
    issues: List[dict] = field(default_factory=list)
    message: str = ""

    def to_dict(self):
        return {
            "status": self.status,
            "issues": self.issues,
            "message": self.message
        }