#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""使用Mock LLM测试完整流程"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 步骤1: 创建Mock LLM客户端
class MockLLMClient:
    """模拟LLM客户端 - 返回预设的智能响应"""

    def __init__(self):
        self.model = "MiniMax-M2.5"

    def invoke(self, messages, **kwargs):
        """模拟LLM调用"""
        # 将消息转换为单字符串
        if isinstance(messages, list):
            content = ""
            for msg in messages:
                if hasattr(msg, 'content'):
                    content = msg.content
                    break
        else:
            content = str(messages)

        # 根据内容生成智能响应
        if "content_planner" in content.lower() or "规划" in content:
            response = '''```json
{
    "slides": [
        {"type": "cover", "title": "云核心网网络架构评估报告", "subtitle": "5G核心网综合评估", "date": "2026年4月"},
        {"type": "toc", "title": "目录", "items": ["一、评估方案介绍", "二、业务背景", "三、整体评估概览", "四至十、各维度评估", "十一、总结与优化建议"]},
        {"type": "intro", "title": "一、评估方案介绍", "content": {"评估维度": "MT0-MT6", "产品域": "5G核心网/物联网/VoNR", "方法": "现场调研+系统检测", "标准": "七维评估体系"}},
        {"type": "background", "title": "二、业务背景与网络概况", "content": {"总体网元数": "235", "容灾组数": "33", "产品域": "7个"}},
        {"type": "overview", "title": "三、整体评估概览", "highlights": [{"维度": "MT0", "通过率": "84%", "风险": "低"}, {"维度": "MT1", "通过率": "73%", "风险": "中"}, {"维度": "MT2", "通过率": "88%", "风险": "低"}, {"维度": "MT3", "通过率": "79%", "风险": "中"}, {"维度": "MT4", "通过率": "60%", "风险": "高"}, {"维度": "MT5", "通过率": "92%", "风险": "低"}, {"维度": "MT6", "通过率": "85%", "风险": "低"}]},
        {"type": "dimension_detail", "title": "四、网元高稳评估(MT0)", "pass_rate": "84%", "risks": ["部分老旧设备需更换"]},
        {"type": "dimension_detail", "title": "五、部署离散度评估(MT1)", "pass_rate": "73%", "risks": ["部分网元未实现同城双活"]},
        {"type": "dimension_detail", "title": "六、组网架构高可用(MT2)", "pass_rate": "88%", "risks": []},
        {"type": "dimension_detail", "title": "七、业务路由高可用(MT3)", "pass_rate": "79%", "risks": ["部分业务路由单一"]},
        {"type": "dimension_detail", "title": "八、网络容量评估(MT4)", "pass_rate": "60%", "risks": ["负载不均衡", "部分网元过载"]},
        {"type": "dimension_detail", "title": "九、网元版本EOS(MT5)", "pass_rate": "92%", "risks": []},
        {"type": "dimension_detail", "title": "十、云平台版本EOS(MT6)", "pass_rate": "85%", "risks": []},
        {"type": "summary", "title": "十一、总结与优化建议", "priorities": [{"priority": "高", "issue": "MT4网络容量", "suggestion": "负载均衡优化"}, {"priority": "高", "issue": "MT1部署离散度", "suggestion": "提升双活覆盖率"}]}
    ]
}
```'''
        elif "data" in content.lower() or "分析" in content:
            response = '''```json
{
    "product_domains": [
        {"name": "5G核心网", "ne_count": 85, "dr_groups": 12},
        {"name": "物联网", "ne_count": 45, "dr_groups": 6},
        {"name": "VoNR", "ne_count": 32, "dr_groups": 4}
    ],
    "overall_results": [
        {"dimension": "MT0", "pass_rate": "84%", "risk": "low"},
        {"dimension": "MT1", "pass_rate": "73%", "risk": "medium"},
        {"dimension": "MT2", "pass_rate": "88%", "risk": "low"}
    ]
}
```'''
        else:
            response = '{"status": "ok"}'

        # 返回响应对象
        return type('Response', (), {'content': response})()


# 步骤2: 注入Mock LLM客户端
print("=" * 60)
print("步骤1: 注入Mock LLM客户端")
print("=" * 60)

from scripts.agents.base import BaseAgent
mock_client = MockLLMClient()
BaseAgent.set_llm_client(mock_client)
print(f"LLM Client injected: {mock_client.model}")

# 步骤3: 运行工作流（启用LLM）
print("\n" + "=" * 60)
print("步骤2: 运行工作流 (use_llm=True)")
print("=" * 60)

from scripts.workflows.agent_manager import generate_ppt_workflow

result = generate_ppt_workflow(
    excel_path='assets/test_operator_A_5G.xlsx',
    user_selections={
        'dimensions': ['MT0', 'MT1', 'MT2', 'MT3', 'MT4', 'MT5', 'MT6'],
        'domains': ['5G核心网'],
        'template': 'default',
        'use_llm': True  # 启用LLM
    }
)

# 步骤4: 输出结果
print("\n" + "=" * 60)
print("结果验证")
print("=" * 60)
print(f"Status: {result['status']}")
if result['status'] == 'success':
    print(f"PPT: {result['data']['ppt_path']}")
    print(f"Slides: {result['data']['slides_count']}")
else:
    print(f"Error: {result.get('message', 'Unknown error')}")