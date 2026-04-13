"""
Excel解析服务
解析云核心网网络架构评估结果Excel文件
"""

import os
import sys
import json
from typing import Dict, List, Any, Optional
from datetime import datetime
from openpyxl import load_workbook

# 动态导入domain实体
try:
    from domain.entities import (
        OntologyModel, ProductDomain, NetworkElement,
        EvaluationItem, EvaluationDimension
    )
except ImportError:
    # 尝试从scripts目录导入
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from scripts.domain.entities import (
        OntologyModel, ProductDomain, NetworkElement,
        EvaluationItem, EvaluationDimension
    )


class ExcelParser:
    """Excel解析器"""

    def __init__(self, file_path: str):
        """初始化解析器"""
        self.file_path = file_path
        self.workbook = None
        self.sheets = {}

    def load(self):
        """加载Excel文件"""
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"文件不存在: {self.file_path}")

        self.workbook = load_workbook(self.file_path)
        self.sheets = {sheet.title: sheet for sheet in self.workbook}
        return self

    def parse_all(self) -> OntologyModel:
        """解析所有Sheet，构建本体模型"""
        # 默认评估维度
        dimensions = [
            EvaluationDimension(code="MT0", name="网元高稳评估", description="评估网元的高可用性和稳定性"),
            EvaluationDimension(code="MT1", name="部署离散度评估", description="评估网元部署的地理分布离散程度"),
            EvaluationDimension(code="MT2", name="组网架构高可用评估", description="评估组网架构的高可用性"),
            EvaluationDimension(code="MT3", name="业务路由高可用评估", description="评估业务路由的高可用性"),
            EvaluationDimension(code="MT4", name="网络容量评估", description="评估网络容量和负载能力"),
            EvaluationDimension(code="MT5", name="网元版本EOS评估", description="评估网元版本生命周期"),
            EvaluationDimension(code="MT6", name="云平台版本EOS评估", description="评估云平台版本生命周期"),
        ]

        # 解析各个Sheet
        product_domains_data = self._parse_all_product_domains()
        data_centers = self._parse_data_centers()
        eval_descriptions = self._parse_evaluation_descriptions()
        overview_results = self._parse_overview_results()  # 解析整体评估概览

        # 转换为领域对象
        product_domains = []
        for pd_data in product_domains_data:
            ne_count = len(pd_data.get("网元列表", []))
            dr_count = len(pd_data.get("容灾组列表", []))
            product_domains.append(ProductDomain(
                name=pd_data.get("名称", ""),
                ne_count=ne_count,
                dr_count=dr_count
            ))

        # 提取网元列表
        network_elements = []
        for pd_data in product_domains_data:
            for ne_data in pd_data.get("网元列表", []):
                network_elements.append(NetworkElement(
                    name=ne_data.get("名称", ""),
                    product_domain=pd_data.get("名称", ""),
                    disaster_recovery_group=""
                ))

        return OntologyModel(
            product_domains=product_domains,
            data_centers=[dc.get("名称", "") for dc in data_centers],
            network_elements=network_elements,
            evaluation_items=[],
            dimensions=dimensions,
            overview_results=overview_results,
            parse_time=datetime.now().isoformat(),
            source_file=self.file_path
        )

    def _parse_all_product_domains(self) -> List[Dict]:
        """解析所有产品域及其数据"""
        product_domains = []

        # 解析MT0-网元高稳评估详情
        mt0_data = self._parse_mt0_high_stability()
        product_domains.extend(mt0_data)

        # 解析MT1-部署离散度评估详情
        mt1_data = self._parse_mt1_deployment_dispersity()
        # 合并到已有产品域
        for pd in product_domains:
            for pd1 in mt1_data:
                if pd["名称"] == pd1["名称"]:
                    pd["容灾组列表"] = pd1.get("容灾组列表", [])
                    pd["部署离散度评估"] = pd1.get("评估项结果", [])

        # 解析MT2-组网架构高可用评估详情
        mt2_data = self._parse_mt2_network_availability()
        for pd in product_domains:
            for pd2 in mt2_data:
                if pd["名称"] == pd2["名称"]:
                    pd["组网架构评估"] = pd2.get("评估项结果", [])

        # 解析MT3-业务路由高可用评估详情
        mt3_data = self._parse_mt3_service_routing()
        for pd in product_domains:
            for pd3 in mt3_data:
                if pd["名称"] == pd3["名称"]:
                    pd["业务路由评估"] = pd3.get("评估项结果", [])

        # 解析MT4-网络容量评估详情
        mt4_data = self._parse_mt4_capacity()
        for pd in product_domains:
            for pd4 in mt4_data:
                if pd["名称"] == pd4["名称"]:
                    pd["网络容量评估"] = pd4.get("容灾组列表", [])

        # 解析网元版本EOX评估
        eox_data = self._parse_mt5_version_eox()
        for pd in product_domains:
            for eox in eox_data:
                if pd["名称"] == eox["产品域"]:
                    pd["版本信息"] = eox.get("网元版本列表", [])

        return product_domains

    def _parse_mt0_high_stability(self) -> List[Dict]:
        """解析MT0-网元高稳评估详情"""
        sheet_name = "MT0-网元高稳评估详情"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        # 提取表头（第一行）
        headers = []
        for cell in next(data):
            headers.append(cell)

        # 跳过空行和说明行，查找实际数据行
        product_domains = {}
        row_idx = 0
        for row in data:
            row_idx += 1
            if not row[0]:  # 跳过空行
                continue
            # 跳过说明行（以"1、"开头）
            if row[0] and isinstance(row[0], str) and row[0].startswith("1、"):
                continue

            product_domain = row[0]  # 产品域
            ne_name = row[1]  # 网元名称
            if not ne_name:
                continue

            # 初始化产品域
            if product_domain not in product_domains:
                product_domains[product_domain] = {
                    "名称": product_domain,
                    "网元列表": []
                }

            # 网元基本信息
            ne = {
                "名称": ne_name,
                "物理类型": row[2] if len(row) > 2 else None,
                "逻辑类型": row[3] if len(row) > 3 else None,
                "高稳评估项结果": []
            }

            # 评估项结果（从第5列开始）
            for i in range(5, len(headers)):
                if headers[i] and i < len(row) and row[i]:
                    ne["高稳评估项结果"].append({
                        "评估项": headers[i],
                        "结果": row[i]
                    })

            product_domains[product_domain]["网元列表"].append(ne)

        return list(product_domains.values())

    def _parse_mt1_deployment_dispersity(self) -> List[Dict]:
        """解析MT1-部署离散度评估详情"""
        sheet_name = "MT1-部署离散度评估详情"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        headers = [cell for cell in next(data)]

        product_domains = {}
        for row in data:
            if not row[0] or not row[1]:
                continue

            product_domain = row[0]
            dr_group = row[1]  # 容灾组
            ne_name = row[2]  # 网元名称

            if product_domain not in product_domains:
                product_domains[product_domain] = {
                    "名称": product_domain,
                    "容灾组列表": []
                }

            # 查找或创建容灾组
            dr = None
            for existing_dr in product_domains[product_domain]["容灾组列表"]:
                if existing_dr["名称"] == dr_group:
                    dr = existing_dr
                    break

            if not dr:
                dr = {"名称": dr_group, "类型": row[2] if len(row) > 2 else None, "包含网元": [], "评估项结果": []}
                product_domains[product_domain]["容灾组列表"].append(dr)

            # 添加网元到容灾组
            if ne_name:
                dr["包含网元"].append({
                    "名称": ne_name,
                    "部署站点": row[5] if len(row) > 5 else None,
                    "部署DC": row[6] if len(row) > 6 else None
                })

            # 评估项结果
            for i in range(8, len(headers)):
                if headers[i] and i < len(row) and row[i]:
                    dr["评估项结果"].append({
                        "评估项": headers[i],
                        "结果": row[i]
                    })

        return list(product_domains.values())

    def _parse_mt2_network_availability(self) -> List[Dict]:
        """解析MT2-组网架构高可用评估详情"""
        sheet_name = "MT2-组网架构高可用评估详情"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        headers = [cell for cell in next(data)]

        product_domains = {}
        for row in data:
            if not row[0] or not row[1]:
                continue

            product_domain = row[0]
            ne_name = row[1]

            if product_domain not in product_domains:
                product_domains[product_domain] = {
                    "名称": product_domain,
                    "评估项结果": []
                }

            # 评估项结果
            for i in range(4, len(headers)):
                if headers[i] and i < len(row) and row[i]:
                    product_domains[product_domain]["评估项结果"].append({
                        "网元": ne_name,
                        "评估项": headers[i],
                        "结果": row[i]
                    })

        return list(product_domains.values())

    def _parse_mt3_service_routing(self) -> List[Dict]:
        """解析MT3-业务路由高可用评估详情"""
        sheet_name = "MT3-业务路由高可用评估详情"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        headers = [cell for cell in next(data)]

        product_domains = {}
        for row in data:
            if not row[0] or not row[1]:
                continue

            product_domain = row[0]
            ne_name = row[1]

            if product_domain not in product_domains:
                product_domains[product_domain] = {
                    "名称": product_domain,
                    "评估项结果": []
                }

            for i in range(4, len(headers)):
                if headers[i] and i < len(row) and row[i]:
                    product_domains[product_domain]["评估项结果"].append({
                        "网元": ne_name,
                        "评估项": headers[i],
                        "结果": row[i]
                    })

        return list(product_domains.values())

    def _parse_mt4_capacity(self) -> List[Dict]:
        """解析MT4-网络容量-网元间均衡度评估详情"""
        sheet_name = "MT4-网络容量-网元间均衡度评估详情"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        headers = [cell for cell in next(data)]

        product_domains = {}
        for row in data:
            if not row[0] or not row[1]:
                continue

            product_domain = row[0]
            dr_group = row[1]  # 容灾组
            metric_name = row[3]  # 指标项
            ne_name = row[4]  # 网元名称

            if product_domain not in product_domains:
                product_domains[product_domain] = {
                    "名称": product_domain,
                    "容灾组列表": []
                }

            # 查找或创建容灾组
            dr = None
            for existing_dr in product_domains[product_domain]["容灾组列表"]:
                if existing_dr["名称"] == dr_group:
                    dr = existing_dr
                    break

            if not dr:
                dr = {"名称": dr_group, "类型": row[2], "指标项列表": []}
                product_domains[product_domain]["容灾组列表"].append(dr)

            # 查找或创建指标项
            metric = None
            for existing_metric in dr["指标项列表"]:
                if existing_metric["名称"] == metric_name:
                    metric = existing_metric
                    break

            if not metric:
                metric = {"名称": metric_name, "网元级分析": [], "容灾组级分析": {}}
                dr["指标项列表"].append(metric)

            # 网元级分析数据
            if ne_name:
                metric["网元级分析"].append({
                    "网元": ne_name,
                    "系统硬件容量": row[5] if len(row) > 5 else None,
                    "系统软件容量": row[6] if len(row) > 6 else None,
                    "当前容量": row[7] if len(row) > 7 else None,
                    "硬件占用率": row[8] if len(row) > 8 else None,
                    "软件占用率": row[9] if len(row) > 9 else None,
                    "硬件负荷偏差": row[10] if len(row) > 10 else None,
                    "软件负荷偏差": row[11] if len(row) > 11 else None
                })

        # 计算容灾组级分析
        for pd in product_domains.values():
            for dr in pd.get("容灾组列表", []):
                for metric in dr.get("指标项列表", []):
                    # 计算平均值
                    hardware_rates = []
                    software_rates = []
                    for ne_data in metric["网元级分析"]:
                        if ne_data.get("硬件占用率"):
                            hardware_rates.append(float(ne_data["硬件占用率"].replace("%", "")))
                        if ne_data.get("软件占用率"):
                            software_rates.append(float(ne_data["软件占用率"].replace("%", "")))

                    if hardware_rates:
                        metric["容灾组级分析"]["平均硬件占用率"] = f"{sum(hardware_rates)/len(hardware_rates):.1f}%"
                    if software_rates:
                        metric["容灾组级分析"]["平均软件占用率"] = f"{sum(software_rates)/len(software_rates):.1f}%"

        return list(product_domains.values())

    def _parse_mt5_version_eox(self) -> List[Dict]:
        """解析网元版本EOX评估详情"""
        sheet_name = "网元版本EOX评估详情"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        product_domains = {}
        for row in data:
            if not row[0] or not row[1]:
                continue

            product_domain = row[0]
            ne_name = row[1]

            if product_domain not in product_domains:
                product_domains[product_domain] = {
                    "产品域": product_domain,
                    "网元版本列表": []
                }

            product_domains[product_domain]["网元版本列表"].append({
                "网元": ne_name,
                "物理类型": row[2] if len(row) > 2 else None,
                "现网版本": row[3] if len(row) > 3 else None,
                "EOS时间": row[4] if len(row) > 4 else None,
                "评估结果": row[5] if len(row) > 5 else None,
                "评估详情": row[6] if len(row) > 6 else None
            })

        return list(product_domains.values())

    def _parse_data_centers(self) -> List[Dict]:
        """解析电信云CloudOS&服务器EOX评估详情"""
        sheet_name = "电信云CloudOS&服务器EOX评估详情"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        # 跳过前两行（标题行）
        data_list = list(data)
        if len(data_list) < 3:
            return []

        dc_list = []
        # 解析DC名称和CloudOS信息
        for i in range(2, len(data_list)):
            row = data_list[i]
            if not row[0]:
                continue

            dc = {
                "名称": row[0],
                "CloudOS": {
                    "版本": row[1] if len(row) > 1 else None,
                    "EOS时间": row[3] if len(row) > 3 else None,
                    "评估结果": row[4] if len(row) > 4 else None
                },
                "服务器": [],
                "评估详情": row[8] if len(row) > 8 else None
            }

            # 解析服务器信息
            for j in range(5, len(data_list)):
                server_row = data_list[j]
                if not server_row[5]:
                    break
                dc["服务器"].append({
                    "类型": server_row[5] if len(server_row) > 5 else None,
                    "数量": server_row[6] if len(server_row) > 6 else None,
                    "EOS时间": server_row[7] if len(server_row) > 7 else None,
                    "评估结果": server_row[8] if len(server_row) > 8 else None
                })

            dc_list.append(dc)

        return dc_list

    def _parse_evaluation_descriptions(self) -> List[Dict]:
        """解析评估项说明"""
        sheet_name = "评估项说明"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        descriptions = []
        for row in data:
            if not row[0]:
                continue

            descriptions.append({
                "评估维度": row[0],
                "产品域": row[1],
                "评估项": row[2],
                "评估说明": row[3],
                "专家经验指导": row[4]
            })

        return descriptions

    def _parse_overview_results(self) -> List[Dict]:
        """解析整体评估概览Sheet"""
        sheet_name = "整体评估概览"
        if sheet_name not in self.sheets:
            return []

        sheet = self.sheets[sheet_name]
        data = sheet.values

        overview_results = []
        for row in data:
            if not row[0] or row[0] == "评估维度":  # 跳过表头和空行
                continue
            # 跳过汇总行（第一列为"汇总"）
            if row[0] == "汇总":
                continue

            result = {
                "评估维度": row[0] if len(row) > 0 else "",
                "评估项": row[1] if len(row) > 1 else "",
                "通过数": row[2] if len(row) > 2 else 0,
                "待改进数": row[3] if len(row) > 3 else 0,
                "不通过数": row[4] if len(row) > 4 else 0,
                "通过率": row[5] if len(row) > 5 else "",
                "风险等级": row[6] if len(row) > 6 else "",
                "整体评估结论": row[7] if len(row) > 7 else "",
                "主要风险描述": row[8] if len(row) > 8 else "",
                "优化建议": row[9] if len(row) > 9 else ""
            }
            overview_results.append(result)

        return overview_results


def parse_excel(file_path: str, output_path: str = None) -> Dict[str, Any]:
    """解析Excel文件并返回本体模型

    Args:
        file_path: Excel文件路径
        output_path: 输出JSON文件路径（可选）

    Returns:
        解析后的本体模型字典
    """
    parser = ExcelParser(file_path)
    parser.load()
    result = parser.parse_all()

    # 如果指定了输出路径，保存JSON
    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"本体模型已保存到: {output_path}")

    return result


if __name__ == "__main__":
    # 测试
    test_file = "../../CNAE看网结果详情.xlsx"
    result = parse_excel(test_file, "../../data/parsed/test_output.json")
    print(f"解析完成，共 {len(result.get('产品域列表', []))} 个产品域")