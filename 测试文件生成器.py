import openpyxl
from openpyxl import Workbook
import random

def create_test_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"

    # 设置表头
    headers = ["运单号码", "增值费用", "应付金额", "其它信息"]
    ws.append(headers)

    # 费用类型池
    fee_types = ["快递运费", "包装服务", "保价", "超长超重", "同城转寄", "异地转寄", "运费", "包装费"]
    
    # 备注池
    remarks = ["正常件", "需要纸箱", "加急", "贵重物品", "大件", "转寄", "正则匹配测试"]

    print("正在生成100行测试数据...")

    # 生成100行数据
    for i in range(1, 101):
        # 生成运单号 SF1001 - SF1100
        sf_number = f"SF{1000 + i}"
        
        # 随机选择费用类型
        fee_type = random.choice(fee_types)
        
        # 随机生成金额 (5.00 - 200.00)
        amount = round(random.uniform(5.0, 200.0), 2)
        
        # 随机选择备注
        info = random.choice(remarks)
        
        ws.append([sf_number, fee_type, amount, info])

    filename = "测试文件.xlsx"
    wb.save(filename)
    print(f"已生成测试文件: {filename} (共100行数据)")

if __name__ == "__main__":
    create_test_excel()
