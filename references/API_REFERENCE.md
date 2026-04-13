# API 参考

## AutoPPTSkill

主入口类，提供统一的Skill调用接口。

### 初始化

```python
from ai_skill_autoppt import AutoPPTSkill

# 使用默认配置
skill = AutoPPTSkill()

# 指定配置文件
skill = AutoPPTSkill(config_path="/path/to/config.yaml")
```

### 方法

#### generate()

生成PPT报告

```python
result = skill.generate(
    excel_path: str,
    template: str = "default",
    dimensions: list = None,
    domains: list = None,
    use_llm: bool = False,
    output_path: str = None
) -> Dict
```

**参数:**

| 参数 | 类型 | 说明 | 默认值 |
|------|------|------|--------|
| excel_path | str | Excel文件路径 | 必需 |
| template | str | 模板名称 | "default" |
| dimensions | list | 评估维度 | ["MT0"-"MT6"] |
| domains | list | 产品域 | [] |
| use_llm | bool | 使用LLM增强 | False |
| output_path | str | 输出路径 | None |

**返回:**

```python
{
    "status": "success",
    "data": {
        "ppt_path": "path/to/report.pptx",
        "slides_count": 13,
        "review_result": {...}
    },
    "message": "PPT生成成功，共13页"
}
```

#### parse()

解析Excel数据（不生成PPT）

```python
data = skill.parse(excel_path: str) -> Dict
```

#### list_templates()

列出可用模板

```python
templates = skill.list_templates() -> list
```

---

## 工作流函数

### generate_ppt_workflow()

核心工作流函数

```python
from ai_skill_autoppt.workflows.agent_manager import generate_ppt_workflow

result = generate_ppt_workflow(
    excel_path: str,
    user_selections: Dict
) -> Dict
```

**参数:**

| 参数 | 类型 | 说明 |
|------|------|------|
| excel_path | str | Excel文件路径 |
| user_selections | Dict | 用户选择配置 |

**user_selections示例:**

```python
{
    "dimensions": ["MT0", "MT1", "MT2", "MT3", "MT4", "MT5", "MT6"],
    "domains": [],
    "template": "default",
    "use_llm": False
}
```

---

## CLI命令

### generate

```bash
ai-skill-autoppt generate [选项]
```

**选项:**

| 选项 | 说明 | 默认值 |
|------|------|--------|
| --excel, -e | Excel文件路径 | 必需 |
| --output, -o | 输出文件路径 | report.pptx |
| --template, -t | 模板名称 | default |
| --dimensions | 评估维度 | MT0-MT6 |
| --use-llm | 使用LLM增强 | False |

### parse

```bash
ai-skill-autoppt parse --excel <file> [--output <file>]
```

### list-templates

```bash
ai-skill-autoppt list-templates
```

---

## 数据类型

### OntologyModel

本体模型数据类型

```python
{
    "product_domains": [
        {
            "name": "5G核心网",
            "ne_count": 10,
            "dr_count": 5,
            "dimensions": {...}
        }
    ],
    "overview_results": [...],
    "total_ne_count": 100,
    "total_dr_count": 50
}
```

### AgentResult

Agent执行结果

```python
{
    "status": "success",
    "data": {...},
    "message": "执行成功"
}
```
