import xlwings as xw
import os
from PySide6.QtWidgets import QMainWindow, QFileDialog

class ClassCompute():
    def __init__(self, exchangeRate, type):
        super().__init__()
        # 执行步骤 1 为平均单价 2为合计表
        self.type = type
        # 输出计算结果状态
        self.result = False
        # 提示用户输入汇率
        self.exchangeRate = float(exchangeRate)
        # 执行选择文件
        self.chooseFile()
        
    def __bool__(self):
        return self.result
    
    def chooseFile(self):
        
        # 创建主窗口
        window = QMainWindow()

        # 弹出文件夹选择对话框
        self.folder_path = QFileDialog.getExistingDirectory(window, "选择文件夹")

        if not self.folder_path:
            self.result = False
            return
        # 选完文件夹执行计算
        if self.type == 1:
            self.startComputeUnitPrice()
        else:
            self.startComputeTotal()

    # 计算平均单价
    def startComputeUnitPrice(self):

        # 获取文件夹下以'Order-SKU-all'为前缀的xlsx文件名
        file_names = [file for file in os.listdir(self.folder_path) if file.startswith('Order-SKU-all') and file.endswith('.xlsx')]

        # 创建一个空的字典来存储筛选后的数据
        filtered_data = {}

        # 启动Excel应用程序
        app = xw.App()

        for file_name in file_names:
            # 拼接完整文件路径
            file_path = os.path.join(self.folder_path, file_name)
            
            # 打开Excel文件
            wb_sku = xw.Book(file_path)
            
            # 获取第一个工作表
            sheet = wb_sku.sheets[0]
            
            # 获取整个工作表的数据范围
            data_range = sheet.used_range
            data = data_range.value
            
            if data:
                for row in data[1:]:  # 跳过第一行表头
                    order_status = row[8]  # 订单状态所在列索引为8
                    picking_remark = row[61]  # 拣货备注所在列索引为61
                    sku = row[31]  # 店铺SKU所在列索引为31
                    
                    if order_status != "已取消" and (not picking_remark or "fake" not in str(picking_remark).lower()):
                        quantity = float(row[35])  # 数量所在列索引为35
                        revenue = float(row[37])  # 营业额所在列索引为37
                        
                        # 将sku相同的数量、营业额进行累加
                        if sku in filtered_data:
                            filtered_data[sku][0] += quantity
                            filtered_data[sku][1] += revenue
                        else:
                            filtered_data[sku] = [quantity, revenue]
                    
        # 计算单价（营业额除以数量再除以汇率）
        for sku, data in filtered_data.items():
            quantity = data[0]
            revenue = data[1]
            unit_price = revenue / quantity / self.exchangeRate
            data.append(unit_price)

        # 创建新的Excel表格并写入累加后的结果
        wb_new_price = xw.Book()
        sheet_new_price = wb_new_price.sheets[0]
        sheet_new_price.range('A1').value = [['店铺SKU', '数量', '单价', '营业额']]
        row_index = 2
        for sku, data in filtered_data.items():
            sheet_new_price.range(f'A{row_index}').value = [sku, data[0], data[2], data[1]]
            row_index += 1

        wb_new_price.save(self.folder_path + '/平均单价合计.xlsx')
        wb_new_price.close()

        self.result = True
        # 退出Excel应用程序
        app.quit()

    # 计算合计
    def startComputeTotal(self):
        # 启动Excel应用程序
        app = xw.App()

        #  打开Excel文件--获取商品型号和分类
        wb_info = xw.Book(self.folder_path + '/商品信息.xlsx')

        # # 选择要操作的工作表
        sheet_info = wb_info.sheets[0]

        data_range_info = sheet_info.used_range

        info_data = {}
        for row in data_range_info.value[3:]: 
            if row[8] != 'None' and row[9] != 'None' :
                info_data[row[8]] = row[9]

        wb_info.close()

        # 广告费表

        # 打开第一个Excel文件--pid对照表
        wb_pid = xw.Book(self.folder_path + '/shopee_edit_price_stock.xlsx')
        sheet_pid = wb_pid.sheets['Sheet1']

        # 获取整个工作表的数据范围
        data_range_pid = sheet_pid.used_range

        pid_data_info = {}
        for row in data_range_pid.value[1:]: 
                pid = int(row[1])
                sku = row[5]
                if pid not in pid_data_info:
                    pid_data_info[pid] = sku

        wb_pid.close()
        # 遍历数据

        # 打开第二个Excel文件
        wb_ad = xw.Book(self.folder_path + '/Shopee-Ads-Overall-Data.xlsx')
        sheet_ad = wb_ad.sheets[0]

        # 获取整个工作表的数据范围
        data_range_ad = sheet_ad.used_range

        ad_data = {}
        # 将pid相同的广告费进行累加
        for row in data_range_ad.value[1:]: 
                pid = int(row[0]) 
                ad_cost = float(row[1])
            
                if pid in ad_data:
                    ad_data[pid] += ad_cost
                else:
                    ad_data[pid] = ad_cost

        wb_ad.close()

        # 遍历去 pid对照表 找到sku对应的分类名
        class_data = {}
        for pid, ad_cost in ad_data.items():
            if pid in pid_data_info:  # 找sku
                sku = pid_data_info[pid]
                if sku in info_data:  # 找分类
                    className = info_data[sku]
                    if className in class_data: # 判断是否存在，存在就累加
                        class_data[className] += ad_cost
                    else:
                        class_data[className] = ad_cost

        # 根据型号去匹配 分类 取出销售数量、收入、利润
        wb_product = xw.Book(self.folder_path + '/销售毛利统计-按商品统计.xlsx')

        # 选择要操作的工作表
        sheet_product = wb_product.sheets[0]

        data_product = sheet_product.used_range

        filtered_data_product = {}
        for row in data_product.value[3:]:  # 从第3行开始遍历数据
                sku = row[1]
                num = row[3] # 数量
                revenue = row[4]  # 销售收入
                grossProfit = row[10]  # 毛利
                if sku in info_data:
                    className = info_data[sku]
                    if className in filtered_data_product:
                        filtered_data_product[className][0] += num
                        filtered_data_product[className][1] += revenue
                        filtered_data_product[className][2] += grossProfit
                    else:
                        filtered_data_product[className] = [num, revenue, grossProfit]

        # 关闭Excel文件
        wb_product.close()


        total_data = {}
        # 遍历 filtered_data_product 和 class_data
        for key in set(filtered_data_product.keys()).union(class_data.keys()):
            # 初始化 total_data[key] 为一个包含四个元素的列表
            total_data.setdefault(key, [None] * 4)

            if key in class_data:
                total_data[key][0] = class_data[key]  # 设置 class_data 的值
            if key in filtered_data_product:
                total_data[key][1] = filtered_data_product[key][0]  # 设置 filtered_data_product 的值
                total_data[key][2] = filtered_data_product[key][1]
                total_data[key][3] = filtered_data_product[key][2]

        # 创建新的 Excel 工作簿
        wb_new_total = xw.Book()
        sheet_new_total = wb_new_total.sheets[0]

        # 写入标题行
        sheet_new_total.range('A1').value = [['类别', '广告费', '销售数量', '销售收入', '利润', '未扣广告毛利率', '扣广告毛利率', '扣平台毛利率']]

        row_index = 2
        for className, data in total_data.items():
            # 计算所需的值
            ad_cost = (data[0] / self.exchangeRate) if data[0] is not None and data[0] != 0 else 0
            sales_quantity = data[1] if data[1] is not None else 0  # 确保 sales_quantity 不是 None
            sales_revenue = data[2] if data[2] is not None else 0  # 确保 sales_revenue 不是 None
            profit = data[3] if data[3] is not None else 0 
            
            # 计算毛利率
            gross_margin_rate = profit / sales_revenue if sales_revenue else 0
            adjusted_margin_rate = (profit - ad_cost) / sales_revenue if sales_revenue else 0
            adjusted_margin_rate_minus_11 = adjusted_margin_rate - 0.11 if adjusted_margin_rate else 0

            # 写入数据
            sheet_new_total.range(f'A{row_index}').value = [
                className,
                ad_cost,
                sales_quantity,
                sales_revenue,
                profit,
                gross_margin_rate,
                adjusted_margin_rate,
                adjusted_margin_rate_minus_11,
            ]

            # 设置后面三列为百分比格式
            sheet_new_total.range(f'F{row_index}:H{row_index}').number_format = '0.00%'
            row_index += 1
            # 设置列宽
            sheet_new_total.range('A:A').column_width = 30 
            sheet_new_total.range('B:B').column_width = 13 
            sheet_new_total.range('C:C').column_width = 10
            sheet_new_total.range('D:D').column_width = 12 
            sheet_new_total.range('E:E').column_width = 12 
            sheet_new_total.range('F:F').column_width = 15 
            sheet_new_total.range('G:G').column_width = 15
            sheet_new_total.range('H:H').column_width = 15

        # 保存并关闭工作簿
        wb_new_total.save(self.folder_path + '/合计表.xlsx')
        wb_new_total.close()
        
        self.result = True
        # 退出Excel应用程序
        app.quit()

def setupFuc(exchangeRate, type):
    res = ClassCompute(exchangeRate, type)
    return res.result