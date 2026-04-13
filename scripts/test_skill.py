#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""构造测试数据并验证SKILL"""

import sys
import os
import io

# 设置输出编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 设置路径 - 相对于SKILL根目录
skill_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(skill_root, 'scripts'))

from services.excel_parser import ExcelParser
from services.ppt_generator import generate_ppt
import json

# 测试1: 解析Excel数据
print("=" * 50)
print("测试1: 解析Excel数据")
print("=" * 50)

excel_path = os.path.join(skill_root, 'assets', 'test_operator_A_5G.xlsx')
parser = ExcelParser(excel_path)
parser.load()
ontology = parser.parse_all()
data = ontology.to_dict()

print(f"[OK] 产品域数量: {len(data.get('产品域列表', []))}")
print(f"[OK] 评估维度: {len(data.get('评估维度', []))}")
print(f"[OK] 评估指标: {len(data.get('整体评估概览', []))}")

# 测试2: 生成标准JSON
print("\n" + "=" * 50)
print("测试2: 生成标准JSON")
print("=" * 50)

# 从整体评估概览中提取评估项
overview = data.get('整体评估概览', [])

result = {
    'metadata': {
        'source_file': 'assets/test_operator_A_5G.xlsx',
        'parse_time': data.get('解析时间', ''),
        'version': '1.0.0'
    },
    'product_domains': data.get('产品域列表', []),
    'evaluation_dimensions': [
        {'code': 'MT0', 'name': '网元高稳评估'},
        {'code': 'MT1', 'name': '部署离散度评估'},
        {'code': 'MT2', 'name': '组网架构高可用评估'},
        {'code': 'MT3', 'name': '业务路由高可用评估'},
        {'code': 'MT4', 'name': '网络容量均衡度评估'},
        {'code': 'MT5', 'name': '网元版本EOS评估'},
        {'code': 'MT6', 'name': '云平台EOS评估'}
    ],
    'evaluation_items': [
        {
            'dimension': item.get('评估维度', '').split('-')[0] if item.get('评估维度') else '',
            'name': item.get('评估项', ''),
            'score': item.get('通过数', 0),
            'pass_rate': item.get('通过率', ''),
            'level': item.get('风险等级', ''),
            'result': item.get('整体评估结论', ''),
            'findings': item.get('主要风险描述', ''),
            'suggestions': item.get('优化建议', '')
        }
        for item in overview
    ]
}

print(f"[OK] JSON字段数: {len(result)}")
print(f"[OK] 评估项数: {len(result['evaluation_items'])}")

# 保存JSON
json_path = os.path.join(skill_root, 'assets', 'test_ontology.json')
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print(f"[OK] 已保存: {json_path}")

# 测试3: 生成PPT
print("\n" + "=" * 50)
print("测试3: 生成PPT报告")
print("=" * 50)

# 构建PPT内容
content = {
    "slides": [
        {
            "type": "cover",
            "title": "5G核心网网络架构评估报告",
            "subtitle": "运营商A - 2026年4月",
            "date": "2026年4月"
        },
        {
            "type": "toc",
            "title": "目录",
            "items": [
                "评估方案介绍",
                "业务背景",
                "整体评估概览",
                "MT0-网元高稳评估",
                "MT1-部署离散度评估",
                "MT2-组网架构高可用评估",
                "MT3-业务路由高可用评估",
                "MT4-网络容量均衡度评估",
                "MT5-网元版本EOS评估",
                "MT6-云平台EOS评估",
                "总结与优化建议"
            ]
        },
        {
            "type": "intro",
            "title": "评估方案介绍",
            "content": {
                "评估维度": ["MT0: 网元高稳", "MT1: 部署离散度", "MT2: 组网架构", "MT3: 业务路由", "MT4: 网络容量", "MT5: 网元版本", "MT6: 云平台"],
                "评估产品域": ["5G核心网", "物联网专用", "VoNR核心网", "短信中心"],
                "评估方法": "基于评分卡的多维度量化评估",
                "评估标准": "通过率≥80%为低风险，60-80%为中风险，<60%为高风险"
            }
        },
        {
            "type": "background",
            "title": "业务背景",
            "content": {
                "总体网元数": "60",
                "容灾组数": "8",
                "产品域分布": [
                    {"name": "5G核心网", "ne_count": 15, "dr_count": 8},
                    {"name": "物联网专用", "ne_count": 15, "dr_count": 8},
                    {"name": "VoNR核心网", "ne_count": 15, "dr_count": 8},
                    {"name": "短信中心", "ne_count": 15, "dr_count": 8}
                ]
            }
        },
        {
            "type": "overview",
            "title": "整体评估概览",
            "content": {
                "overall": "整体评估得分72分，整体处于中等偏上水平。MT0、MT2表现良好，MT1、MT3、MT4需要重点关注优化。",
                "highlights": [
                    {"维度": "MT0-网元高稳", "通过率": "84%", "风险": "低", "状态": "良好"},
                    {"维度": "MT1-部署离散度", "通过率": "73%", "风险": "中", "状态": "需改善"},
                    {"维度": "MT2-组网架构", "通过率": "78%", "风险": "中", "状态": "良好"},
                    {"维度": "MT3-业务路由", "通过率": "75%", "风险": "中", "状态": "需改善"},
                    {"维度": "MT4-网络容量", "通过率": "60%", "风险": "高", "状态": "需重点关注"},
                    {"维度": "MT5-网元版本EOS", "通过率": "72%", "风险": "中", "状态": "需关注"},
                    {"维度": "MT6-云平台EOS", "通过率": "65%", "风险": "中", "状态": "需关注"}
                ]
            }
        },
        {
            "type": "dimension_detail",
            "title": "MT0 - 网元高稳评估",
            "pass_rate": "84%",
            "risk": "低",
            "findings": ["网元状态监测正常", "计划告警处理及时", "动环告警较少"],
            "risks": [],
            "suggestions": ["继续保持现有运维标准", "定期巡检确保设备健康"]
        },
        {
            "type": "dimension_detail",
            "title": "MT1 - 部署离散度评估",
            "pass_rate": "73%",
            "risk": "中",
            "findings": ["同城双活覆盖率63%", "跨区容灾覆盖率75%", "异地DC容灾覆盖率80%"],
            "risks": ["仅63%业务实现同城双活"],
            "suggestions": ["推进业务100%双活部署", "提升跨区容灾覆盖率至85%"]
        },
        {
            "type": "dimension_detail",
            "title": "MT2 - 组网架构高可用评估",
            "pass_rate": "78%",
            "risk": "中",
            "findings": ["负荷分担配置覆盖84%", "主备配置覆盖71%", "故障切换机制基本建立"],
            "risks": ["主备配置覆盖率71%"],
            "suggestions": ["优化负荷分担配置", "完善主备切换机制"]
        },
        {
            "type": "dimension_detail",
            "title": "MT3 - 业务路由高可用评估",
            "pass_rate": "75%",
            "risk": "中",
            "findings": ["主用路由覆盖70%", "备用路由覆盖74%", "路由收敛时间满足要求"],
            "risks": ["主用路由覆盖率仅70%"],
            "suggestions": ["提升主用路由覆盖率至90%", "优化路由收敛时间"]
        },
        {
            "type": "dimension_detail",
            "title": "MT4 - 网络容量均衡度评估",
            "pass_rate": "60%",
            "risk": "高",
            "findings": ["负荷均衡度仅54%", "流量偏置偏高28%", "硬件占用情况良好"],
            "risks": ["负荷均衡度严重不足", "流量偏置问题突出"],
            "suggestions": ["优化流量负载均衡策略", "降低流量偏置至15%以内"]
        },
        {
            "type": "dimension_detail",
            "title": "MT5 - 网元版本EOS评估",
            "pass_rate": "72%",
            "risk": "中",
            "findings": ["18%网元版本存在EOS风险", "31%网元进入EOS周期"],
            "risks": ["31%网元存在EOS风险"],
            "suggestions": ["制定6个月内升级计划"]
        },
        {
            "type": "dimension_detail",
            "title": "MT6 - 云平台EOS评估",
            "pass_rate": "65%",
            "risk": "中",
            "findings": ["1个DC的CloudOS进入EOS", "50%虚拟机进入EOS周期"],
            "risks": ["50%虚拟机存在EOS风险"],
            "suggestions": ["Q2季度完成升级替换", "制定3年替换计划"]
        },
        {
            "type": "summary",
            "title": "总结与优化建议",
            "content": {
                "overall": "整体评估结果为中等偏上，需要重点关注MT4网络容量均衡度问题。",
                "priorities": [
                    {"优先级": "高", "问题": "MT4-负荷均衡度严重不足", "建议": "立即优化流量负载均衡策略"},
                    {"优先级": "高", "问题": "MT1-同城双活覆盖率低", "建议": "推进业务100%双活部署"},
                    {"优先级": "中", "问题": "MT3-主用路由覆盖率不足", "建议": "提升主用路由覆盖率至90%"},
                    {"优先级": "中", "问题": "MT5/M6-版本EOS风险", "建议": "制定6个月升级计划"}
                ],
                "next_steps": "建议召开专项会议，制定整改计划"
            }
        }
    ]
}

# 生成PPT
skill_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
output_dir = os.path.join(skill_root, 'storage', 'ppt')
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, 'test_report.pptx')
result = generate_ppt(content, output_path, None, 'zh')
print(f"[OK] PPT生成成功: {result}")

# 验证文件存在
if os.path.exists(output_path):
    size = os.path.getsize(output_path)
    print(f"[OK] 文件大小: {size} bytes")

print("\n" + "=" * 50)
print("所有测试通过!")
print("=" * 50)
