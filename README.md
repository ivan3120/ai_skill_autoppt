# AI Skill: 云核心网评估报告生成

智能化生成云核心网网络架构评估PPT报告的AI Skill。

## 功能特性

- 📊 **智能数据解析** - 自动解析Excel评估数据
- 📝 **结构规划** - 智能规划PPT报告结构
- 🎨 **风格学习** - 提取和应用PPT模板风格
- 🖼️ **图片匹配** - 自动匹配相关网络拓扑图
- ✅ **质量审核** - 自动审核PPT内容质量
- 🔄 **修改处理** - 支持自然语言修改

## 快速开始

### 安装

```bash
pip install ai-skill-autoppt
```

### 使用

#### CLI
```bash
ai-skill-autoppt generate --excel data.xlsx --output report.pptx
```

#### Python
```python
from ai_skill_autoppt import AutoPPTSkill

skill = AutoPPTSkill()
result = skill.generate(excel_path="data.xlsx")
print(result)
```

#### 自然语言
在Claude Code中：
> "帮我分析这份5G核心网评估数据并生成PPT报告"

## 项目结构

```
ai_skill_autoppt/
├── src/ai_skill_autoppt/   # 源代码
├── prompts/                # LLM提示词
├── skills/                 # Skill定义
├── resources/              # 资源文件
│   ├── images/            # 配图
│   ├── templates/         # PPT模板
│   ├── data/              # 示例数据
│   └── knowledge/         # 背景知识
├── scripts/                # 工具脚本
└── docs/                   # 文档
```

## 文档

- [用户指南](docs/USER_GUIDE.md)
- [API参考](docs/API_REFERENCE.md)
- [故障排查](docs/TROUBLESHOOTING.md)

## License

MIT
