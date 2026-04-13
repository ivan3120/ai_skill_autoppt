# AI Skill: 云核心网网络架构评估报告生成

## 概述

**Skill ID**: ai_skill_autoppt
**版本**: v1.1.0
**功能**: 智能化生成云核心网网络架构评估PPT报告
**适用场景**: 5G核心网、物联网、VoNR等网络架构的评估报告自动化生成

**设计原则**:
- 本Skill无LLM配置依赖，内置规则引擎可独立运行
- 宿主Agent可选择注入LLM客户端以增强能力
- 模板风格完整继承，自动提取配色、字体、表格、图表等视觉元素

## 能力描述

本Skill能够：
1. **解析评估数据** - 解析Excel格式的云核心网评估数据
2. **构建本体模型** - 提取产品域、网元、容灾组等核心要素
3. **规划PPT结构** - 基于数据内容智能规划报告结构
4. **设计视觉风格** - 提取和应用PPT模板风格（配色、字体、表格、图表）
5. **匹配相关图片** - 匹配网络拓扑图、原理图等素材
6. **生成PPT报告** - 自动生成专业的评估报告PPT
7. **质量审核** - 自动审核PPT内容质量

## Agent能力 - 自然语言描述

### DataAnalyst（数据分析师）

**像一位资深的数据分析师，从Excel表格中提取网络评估的核心信息**

- **做什么**: 读取Excel文件，解析出产品域、网元、容灾组等评估对象，构建数据模型
- **输入**:
  ```json
  {"excel_path": "/data/5G评估表.xlsx"}
  ```
- **输出**:
  ```json
  {
    "product_domains": [
      {"name": "5G核心网", "ne_count": 150, "dr_groups": 20},
      {"name": "物联网", "ne_count": 80, "dr_groups": 10}
    ],
    "dimensions": ["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"],
    "overall_results": [
      {"评估维度": "MT0-网元高稳", "通过率": "85%", "风险等级": "低"},
      {"评估维度": "MT1-部署离散度", "通过率": "60%", "风险等级": "中"}
    ]
  }
  ```
- **示例**: 用户说"帮我分析这份5G核心网评估数据" → Agent解析出4个产品域、200+网元

---

### ContentPlanner（内容规划师）

**像一位专业的PPT策划师，为评估报告设计清晰的故事线和内容结构**

- **做什么**: 根据解析后的数据，规划PPT报告的页面结构、故事线逻辑
- **输入**:
  ```json
  {
    "selected_dimensions": ["MT0", "MT1", "MT2"],
    "selected_domains": ["5G核心网"],
    "parsed_data": {...}
  }
  ```
- **输出**:
  ```json
  {
    "slides": [
      {"type": "cover", "title": "云核心网网络架构评估报告"},
      {"type": "toc", "title": "目录", "items": ["一、评估方案介绍", "二、业务背景"]},
      {"type": "overview", "title": "整体评估概览", "content": {...}},
      {"type": "dimension_detail", "title": "一、网元高稳评估(MT0)", ...},
      {"type": "summary", "title": "总结与优化建议", ...}
    ]
  }
  ```
- **示例**: 用户说"生成一份评估报告" → Agent自动规划13页完整报告结构

---

### VisualDesigner（视觉设计师）

**像一位专业设计师，学习PPT模板的视觉风格并应用到新PPT中**

- **做什么**: 分析模板的配色、字体、布局、表格样式、图表样式，将风格配置提取为可复用参数
- **输入**:
  ```json
  {"template_path": "/templates/客户汇报.pptx"}
  ```
- **输出**:
  ```json
  {
    "style_config": {
      "配色": {
        "主色": "#1E3A5F",
        "辅色": "#F5F5F5",
        "强调色": "#FF6B00",
        "文字色": "#333333"
      },
      "字体": {
        "主标题": "微软雅黑 Bold 32pt",
        "副标题": "微软雅黑 Bold 24pt",
        "正文": "微软雅黑 18pt"
      },
      "表格样式": {"title_bg": "#1E3A5F", "border": true},
      "图表样式": {"type": "COLUMN_CLUSTERED", "colors": ["#1E3A5F", "#FF6B00"]}
    }
  }
  ```
- **示例**: 用户说"用这个模板风格生成报告" → Agent提取"深蓝色主题、微软雅黑字体、橙色强调色"

---

### ImageMatcher（图片匹配师）

**像一位图片库管理员，从知识库中匹配合适的原理性介绍图片**

- **做什么**: 根据评估维度从图片库中匹配对应的网络拓扑图、原理图
- **输入**:
  ```json
  {
    "slides": [...],
    "dimension": "MT0",
    "item": "网元高稳",
    "product_domain": "5G核心网"
  }
  ```
- **输出**:
  ```json
  {
    "image_mapping": {
      3: "/images/zh/MT0/网元高稳.png",
      5: "/images/zh/MT1/部署离散度.png"
    }
  }
  ```
- **示例**: 用户说"在MT0页面配一张网元架构图" → Agent匹配到对应的网络拓扑图

---

### QualityReviewer（质量审核员）

**像一位专业的PPT审核专家，检查报告的内容质量和格式规范**

- **做什么**: 审核PPT报告的内容准确性、逻辑连贯性、格式规范性
- **输入**:
  ```json
  {
    "ppt_content": [...],
    "original_data": {...},
    "style_config": {...}
  }
  ```
- **输出**:
  ```json
  {
    "status": "passed",
    "score": 95,
    "issues": [],
    "strengths": ["结构完整", "格式规范", "数据准确"]
  }
  ```
- **示例**: 用户说"检查一下这份报告" → Agent发现"缺少封面页"或"整体审核通过"

---

### ModificationHandler（修改处理器）

**像一位贴心的助手，解析用户的自然语言修改指令并执行**

- **做什么**: 解析用户的自然语言修改指令，如"把标题改为..."、"在第3页添加图表..."
- **输入**:
  ```json
  {
    "modification_text": "把封面标题改为'5G核心网评估报告'",
    "current_content": [...]
  }
  ```
- **输出**:
  ```json
  {
    "status": "success",
    "modifications": [
      {"page": 0, "operation": "modify_title", "old_value": "云核心网", "new_value": "5G核心网评估报告"}
    ]
  }
  ```
- **示例**: 用户说"把第一页标题改成'2024年度评估报告'" → Agent自动修改标题内容

---

## 调用方式

### 自然语言调用
用户可以说：
- "帮我分析这份5G核心网评估数据并生成PPT报告"
- "使用ai_skill_autoppt解析Excel文件，生成网络架构评估报告"
- "生成云核心网评估PPT"

### Python API调用
```python
# 方式1: 免注入LLM - 纯规则引擎运行
from ai_skill_autoppt import AutoPPTSkill

skill = AutoPPTSkill()
result = skill.generate(excel_path="path/to/data.xlsx")
print(f"PPT生成成功: {result['data']['ppt_path']}")

# 方式2: 注入LLM - 智能增强模式
from scripts.agents.base import BaseAgent
BaseAgent.set_llm_client(host_llm_client)  # 注入宿主Agent的LLM客户端

skill = AutoPPTSkill()
result = skill.generate(excel_path="path/to/data.xlsx")
```

### CLI调用
```bash
ai-skill-autoppt generate --excel path/to/data.xlsx --output report.pptx

# 指定模板
ai-skill-autoppt generate --excel data.xlsx --template custom --output report.pptx
```

## 输入格式

### 必需输入
```json
{
  "excel_path": "string - Excel评估数据文件路径"
}
```

### 可选输入
```json
{
  "template": "string - PPT模板名称 (default/custom)",
  "dimensions": ["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"],
  "domains": ["5G核心网", "物联网"]
}
```

## 输出格式

```json
{
  "status": "success/error",
  "data": {
    "ppt_path": "string - 生成PPT的路径",
    "slides_count": "number - 幻灯片数量",
    "review_result": "object - 质量审核结果"
  },
  "message": "string - 执行消息"
}
```

## 评估维度

本Skill支持以下评估维度：

| 维度 | 说明 |
|------|------|
| MT0 | 网元高稳评估 |
| MT1 | 部署离散度评估 |
| MT2 | 组网架构高可用评估 |
| MT3 | 业务路由高可用评估 |
| MT4 | 网络容量-网元间均衡度评估 |
| MT5 | 网元版本EOS评估 |
| MT6 | 云平台版本EOS评估 |

## 资源文件

- **templates/** - PPT模板文件
- **images/** - 网络拓扑图、原理图等配图
- **references/** - 5G核心网评估相关背景知识

## 约束条件

1. Excel文件必须包含"整体评估概览"Sheet
2. 支持的格式：.xlsx, .xls
3. 模板文件格式：.pptx, .ppt

## 示例数据

详见 `assets/test_operator_A_5G.xlsx`

---

**安装方式**:
```bash
pip install ai-skill-autoppt
```
