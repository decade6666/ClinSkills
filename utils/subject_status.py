# ============================================
# 临床试验数据处理 - Notebook 一体化版本
# 直接在 Jupyter Notebook 中运行此单元格
# ============================================

import pandas as pd
import yaml
import warnings
from typing import Dict, List, Callable, Any

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class SubjectStatus:
    """通用临床试验数据处理器"""
    
    def __init__(self, config_path: str, raw_path: str = None, output_path: str = None):
        """
        初始化处理器
        
        Args:
            config_path: 配置文件路径
            raw_path: 原始数据文件路径（可选，如不提供则使用配置文件中的路径）
            output_path: 输出文件路径（可选，如不提供则使用配置文件中的路径）
        """
        with open(config_path, "r", encoding="utf-8") as f:
            self.cfg = yaml.safe_load(f)
        
        # 优先使用参数传入的路径，否则使用配置文件中的路径
        self.raw_path = raw_path if raw_path is not None else self.cfg["path"]["raw_path"]
        self.output_path = output_path if output_path is not None else self.cfg["path"]["output_path"]
        
        self.index_cols = self.cfg["data"]["index_columns"]
        self.sheets_config = self.cfg["data"]["sheets"]
        self.status_rules = self.cfg["rules"]["status_judgment"]
        self.status_order = self.cfg["output"]["status_order"]
    
    def load_sheet(self, sheet_name: str, columns: List[str], 
                   rename_map: Dict[str, str] = None) -> pd.DataFrame:
        """加载Excel中的单个sheet"""
        df = pd.read_excel(
            self.raw_path, 
            sheet_name, 
            usecols=columns, 
            header=1, 
            dtype=str
        )
        
        if rename_map:
            df = df.rename(columns=rename_map)
        
        return df
    
    def load_all_data(self) -> pd.DataFrame:
        """加载并合并所有数据表"""
        base_config = self.sheets_config[0]
        final = self.load_sheet(
            base_config["sheet_name"],
            self.index_cols,
            base_config.get("rename_map")
        )
        
        for sheet_config in self.sheets_config[1:]:
            columns = self.index_cols + sheet_config["data_columns"]
            df = self.load_sheet(
                sheet_config["sheet_name"],
                columns,
                sheet_config.get("rename_map")
            )
            final = final.merge(df, on=self.index_cols, how="left")
        
        return final
    
    def create_judgment_function(self) -> Callable:
        """根据配置动态创建判断函数"""
        rules = self.status_rules
        
        def judge(row: pd.Series) -> str:
            """动态生成的判断函数"""
            for rule in rules:
                conditions = rule["conditions"]
                result = rule["result"]
                
                all_conditions_met = True
                for condition in conditions:
                    col = condition["column"]
                    op = condition["operator"]
                    value = condition.get("value")
                    
                    if op == "is_not_empty":
                        if row[col] == "" or pd.isna(row[col]):
                            all_conditions_met = False
                            break
                    elif op == "is_empty":
                        if row[col] != "" and not pd.isna(row[col]):
                            all_conditions_met = False
                            break
                    elif op == "equals":
                        if row[col] != value:
                            all_conditions_met = False
                            break
                    elif op == "not_equals":
                        if row[col] == value:
                            all_conditions_met = False
                            break
                
                if all_conditions_met:
                    return result
            
            return rules[-1].get("default", "其他")
        
        return judge
    
    def process(self) -> pd.DataFrame:
        """执行完整的数据处理流程"""
        print("正在加载数据...")
        final = self.load_all_data()
        
        print("正在应用判断规则...")
        judge_func = self.create_judgment_function()
        final["受试者状态"] = final.apply(judge_func, axis=1)
        
        print("正在生成统计结果...")
        group_by_cols = self.cfg["output"]["group_by"]
        result = (
            final
            .groupby(group_by_cols + ["受试者状态"])
            .size()
            .unstack(fill_value=0)
            .reindex(columns=self.status_order, fill_value=0)
            .reset_index()
        )
        
        return result
    
    def save_result(self, result: pd.DataFrame):
        """保存结果到Excel"""
        result.to_excel(self.output_path, index=False)
        print(f"结果已保存至: {self.output_path}")


# ============================================
# 执行处理
# ============================================

if __name__ == "__main__":
    # 方式1: 使用配置文件中的路径
    processor = SubjectStatus("config.yaml")
    
    # 方式2: 通过参数指定路径
    # processor = SubjectStatus(
    #     config_path="config.yaml",
    #     raw_path="data/原始数据.xlsx",
    #     output_path="output/统计结果.xlsx"
    # )
    
    # 执行处理
    result = processor.process()
    
    # 显示结果
    print("\n统计结果:")
    print(result)
    
    # 保存结果
    processor.save_result(result)

# 如果在 notebook 中，也可以直接调用
# processor = SubjectStatus("config.yaml", raw_path="your_data.xlsx", output_path="result.xlsx")
# result = processor.process()
# display(result)
# processor.save_result(result)