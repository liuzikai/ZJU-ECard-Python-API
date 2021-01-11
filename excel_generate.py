import xlwt
import datetime
from termcolor import colored, cprint

work_book = xlwt.Workbook()
work_sheet = work_book.add_sheet("general")
column_count = 0


def generate_excel(records, workbook_filename):

    def write_row(column):
        # print(column)
        global column_count
        for i in range(0, len(column)):
            if column[i] != "":
                work_sheet.write(column_count, i, column[i])
        column_count += 1

    write_row(["交易类型", "日期", "分类", "子分类", "账户1", "账户2", "金额", "成员", "商家", "项目", "备注"])

    for record in records:

        # record: [Time, Type, Subtype, Value, Remainder]
        time = record[0]
        main_type = record[1]
        sub_type = record[2]
        value = float(record[3])
        remainder = record[4]

        _meals = ["海宁食堂一楼", "紫金港校区休闲餐厅", "紫金港校区风味餐厅", "玉泉一食堂一楼", "文科组团一楼食堂"]
        _snacks = ["海宁市华联购中心有限公司", "新宇物业集团有限公司玉泉宿舍水控", "师生交流吧服务部"]

        if main_type in _meals:
            write_row(["支出", time, "食品酒水", "早午晚餐", "ZJU Card", "", -value])
        elif main_type in _snacks:
            write_row(["支出", time, "食品酒水", "零食饮品", "ZJU Card", "", -value])
        elif main_type == "海宁国际校区图书信息中心自助打印复印":
            write_row(["支出", time, "学习进修", "文印服务", "ZJU Card", "", -value])
        elif sub_type == "银行转账":
            write_row(["转账", time, "", "", "BOC", "ZJU Card", value])
        elif sub_type == "支付宝转账":
            print(colored("[Selection Required] 支付宝转账", "yellow"))
            selection = int(input("请确认转出账户：0: BOC, 1: Alipay "))
            write_row(["转账", time, "", "", ["BOC", "Alipay"][selection], "ZJU Card", value])
        else:
            print(colored("[Unexpected Record] Omitted", 'red'))
            print(time, main_type, sub_type, value, remainder)

    work_book.save(workbook_filename)


if __name__ == "__main__":
    # Unit test

    test_time = datetime.date.today().strftime("%Y-%m-%d") + " 12:00:00"
    records = [
        [test_time, "", "银行转账", "100", "100.00"],
        [test_time, "海宁食堂一楼", "持卡人消费", "-10", "90.00"],
        [test_time, "紫金港校区休闲餐厅", "校园宝消费", "-10", "80.00"],
        [test_time, "海宁市华联购中心有限公司", "持卡人消费", "-10", "70.00"],
        [test_time, "海宁国际校区图书信息中心自助打印复印", "持卡人消费", "-10", "60.00"],
        [test_time, "", "支付宝转账", "100", "160.00"],
        [test_time, "main_type", "sub_type", "-10", "150.00"]
    ]

    generate_excel(records, "test.xls")
    cprint("[Excel Generate] Success", 'blue')
