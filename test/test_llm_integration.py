#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""测试LLM集成 - 模拟宿主Agent注入LLM客户端"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 模拟宿主Agent注入LLM客户端
class MockLLMClient:
    """模拟LLM客户端 - 实际使用时替换为真实的LLM客户端"""

    def __init__(self):
        self.model = "Mock-LLM"
        print("[Mock LLM] Client initialized")

    def invoke(self, messages, **kwargs):
        """模拟LLM调用 - 支持LangChain消息格式"""
        # 处理LangChain消息格式 (list of messages)
        if isinstance(messages, list):
            # 提取用户消息内容
            user_content = ""
            for msg in messages:
                if hasattr(msg, 'content'):
                    user_content = msg.content
                    break
                elif isinstance(msg, dict):
                    user_content = msg.get('content', '')
                    break
            prompt_len = len(user_content)
        else:
            prompt_len = len(str(messages))
            user_content = str(messages)

        print(f"[Mock LLM] Received prompt length: {prompt_len} chars")

        # 根据prompt内容返回不同的模拟响应
        if "5G" in user_content or "核心网" in user_content:
            return type('MockResponse', (), {
                'content': '''```json
{
    "analysis": "5G核心网网络架构整体评估结果良好，产品域包括5G核心网、物联网专用、VoNR核心网等7个产品域，共60+网元。",
    "key_findings": [
        "MT0-网元高稳评估通过率84%，风险等级低",
        "MT1-部署离散度评估通过率73%，风险等级中",
        "MT4-网络容量均衡度评估通过率60%，风险等级高，需重点关注"
    ],
    "recommendations": [
        "建议优先优化MT4网络容量负载均衡",
        "提升MT1同城双活覆盖率至85%以上",
        "制定网元版本EOS升级计划"
    ]
}
```'''
            })()

        # 默认响应
        return type('MockResponse', (), {
            'content': '''```json
{
    "analysis": "数据解析完成",
    "key_findings": [],
    "recommendations": []
}
```'''
        })()


# 步骤1: 注入LLM客户端
print("=" * 60)
print("步骤1: 模拟宿主Agent注入LLM客户端")
print("=" * 60)

from scripts.agents.base import BaseAgent

mock_client = MockLLMClient()
BaseAgent.set_llm_client(mock_client)

# 验证注入成功
print(f"BaseAgent._llm_client: {BaseAgent._llm_client}")

# 创建Agent实例验证
from scripts.agents.data_analyst import DataAnalystAgent
agent = DataAnalystAgent()
print(f"_has_llm(): {agent._has_llm()}")

# 步骤2: 运行工作流（启用LLM）
print("\n" + "=" * 60)
print("步骤2: 运行工作流（use_llm=True）")
print("=" * 60)

from scripts.workflows.agent_manager import generate_ppt_workflow

result = generate_ppt_workflow(
    excel_path='assets/test_operator_A_5G.xlsx',
    user_selections={
        'dimensions': ['MT0', 'MT1', 'MT2'],
        'domains': ['5G核心网'],
        'template': 'default',
        'use_llm': True  # 启用LLM
    }
)

print("\n" + "=" * 60)
print("步骤3: 结果验证")
print("=" * 60)
print(f"Status: {result['status']}")
if result['status'] == 'success':
    print(f"PPT: {result['data']['ppt_path']}")
    print(f"Slides: {result['data']['slides_count']}")
else:
    print(f"Error: {result.get('message', 'Unknown error')}")