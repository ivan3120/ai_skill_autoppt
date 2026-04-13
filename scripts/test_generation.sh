#!/bin/bash
# 测试脚本 - 验证PPT生成功能

set -e

echo "🧪 正在测试 ai_skill_autoppt..."
echo ""

# 测试1: 解析Excel
echo "📊 测试1: 解析Excel数据"
python -c "
import sys
sys.path.insert(0, 'src')
from ai_skill_autoppt.services.excel_parser import ExcelParser

parser = ExcelParser('resources/data/test_operator_A_5G.xlsx')
parser.load()
ontology = parser.parse_all()
print(f'✅ 解析成功: {len(ontology.product_domains)} 个产品域')
"

# 测试2: 列出模板
echo ""
echo "🎨 测试2: 列出模板"
python -c "
import sys
sys.path.insert(0, 'src')
from ai_skill_autoppt import AutoPPTSkill

skill = AutoPPTSkill()
templates = skill.list_templates()
print(f'✅ 找到 {len(templates)} 个模板')
for t in templates:
    print(f'   - {t}')
"

# 测试3: 生成PPT (可选)
if [ "$1" = "--full" ]; then
    echo ""
    echo "📝 测试3: 生成完整PPT"
    python -c "
import sys
sys.path.insert(0, 'src')
from ai_skill_autoppt import AutoPPTSkill

skill = AutoPPTSkill()
result = skill.generate(
    excel_path='resources/data/test_operator_A_5G.xlsx',
    template='default',
    dimensions=['MT0', 'MT1', 'MT2', 'MT3', 'MT4', 'MT5', 'MT6']
)

if result.get('status') == 'success':
    print(f'✅ 生成成功: {result[\"data\"][\"ppt_path\"]}')
    print(f'   页数: {result[\"data\"][\"slides_count\"]}')
else:
    print(f'❌ 生成失败: {result.get(\"message\")}')
"
fi

echo ""
echo "✅ 测试完成!"
