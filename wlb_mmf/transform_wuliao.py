import datetime
import xlrd
import os
import xlwt
from xlutils.copy import copy
import logging
logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s %(filename)s %(funcName)s [line:%(lineno)d] %(levelname)s %(message)s')

# 设置屏幕打印的格式
sh = logging.StreamHandler()
sh.setFormatter(formatter)
logger.addHandler(sh)

# 设置log保存
fh = logging.FileHandler("test.log", encoding='utf8')
fh.setFormatter(formatter)
logger.addHandler(fh)

"""
    打包命令
    pyinstaller -F -n "一键运行" -i C:\\Users\zsl\Desktop\图标1转.ico transform_wuliao.py
    pyinstaller -F 一键运行.spec
    
"""

def main():
    trans_obj = Transform()
    trans_obj.write_to_excel()
    trans_obj.deal_info()


class Transform():
    def __init__(self):
        self.template_name = '模板文件\物料汇总模板.xlsx'
        self.table_name_list = self.get_all_template_data()
        self.all_xlsx_list = self.get_all_xlsx()
        self.tody = datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
        self.row_indext = 1
        self.dateFormat = xlwt.XFStyle()
        self.dateFormat.num_format_str = 'yyyy/mm/dd'
        self.error_sheet_list = []
        self.empty_sheet_list = []
        self.xlsx_pwd_split = ""

        if os.path.exists(f"工作簿{self.tody}.xlsx"):
            os.remove(f"工作簿{self.tody}.xlsx")

    def get_all_template_data(self):
        """
        获取模板文件内所有地区中英文字段
        :return:
        """
        logging.info("从模板文件中获取字段。。。。")
        data = xlrd.open_workbook(self.template_name)
        return data.sheets()[0].row_values(0)


    def get_all_xlsx(self):
        logging.info("扫描当前目录下所有.xlsx文件。。。")
        all_xlsx_list = []
        for r,d,f in os.walk(os.getcwd() + "\物料表"):
            for name in f:
                if name.endswith(".xlsx") and not name.startswith("~$"):
                    all_xlsx_list.append(os.path.join(r,name))
        all_xlsx_list_str = '\n'.join(all_xlsx_list)
        logging.info(f"共找到以下【{len(all_xlsx_list)}】个xlsx文件：\n{all_xlsx_list_str}")
        logging.info("*"*100)
        return all_xlsx_list

    def write_to_excel(self):
        if not self.all_xlsx_list:
            raise Exception("没有找到待整理文件，请确认")
        work = xlrd.open_workbook(self.template_name)
        sh = work.sheet_by_index(0)
        self.old_content = copy(work)
        self.ws = self.old_content.get_sheet(0)
        for xlsx_pwd in self.all_xlsx_list:
            try:
                # 用来打印相对路径
                self.xlsx_pwd_split = xlsx_pwd.split(os.getcwd())[-1]
                self.cal_xlsx_data(xlsx_pwd)
            except Exception as e:
                self.error_sheet_list.append([str(e),self.xlsx_pwd_split,""])
                logging.error(e)
                if os.path.exists(f"工作簿{self.tody}.xlsx"):
                    os.remove(f"工作簿{self.tody}.xlsx")
                raise Exception(e)
        logging.info(f"完成填写，即将退出程序。。。")


    def cal_xlsx_data(self,xlsx_pwd):
        logging.info(f"打开文件：【{xlsx_pwd}】")
        data = xlrd.open_workbook(xlsx_pwd)
        for sheet_data in data.sheets():
            logging.info(f"打开sheet页【{sheet_data.name}】")

            if sheet_data.nrows < 2 :
                logging.debug(f"文件：【{self.xlsx_pwd_split}】，sheet:【{sheet_data.name}】，空表，默认忽略该sheet页")
                self.empty_sheet_list.append([self.xlsx_pwd_split,sheet_data.name])
                continue

            if "机加件" in sheet_data.name or "机加件" in xlsx_pwd.split("\\")[-1]:
                # 忽略 机加件
                logging.debug(f"文件：【{self.xlsx_pwd_split}】，sheet:【{sheet_data.name}】，默认忽略该sheet页")
                continue

            sheet_tbody_list = sheet_data.row_values(1)
            if sheet_tbody_list.count("入库日期") < 1 and sheet_tbody_list.count("物料状态") < 1:
                error_msg = f"文件：【{self.xlsx_pwd_split}】，sheet:【{sheet_data.name}】，该表不存在【入库日期】和【物料状态】列，默认忽略该sheet页"
                logging.debug(error_msg)
                self.error_sheet_list.append(["该表不存在【入库日期】和【物料状态】列：","文件地址："+self.xlsx_pwd_split, "sheet页名称："+sheet_data.name])
                continue

            # 校验表头
            self.verify_tbody(sheet_tbody_list,sheet_data.name)

            # 填写内容
            self.write_data(sheet_tbody_list,sheet_data)



    def write_data(self,sheet_tbody_list,sheet_data):
        rkrq_index = sheet_tbody_list.index("入库日期")
        wlzt_index = sheet_tbody_list.index("物料状态")
        for row_index in range(2,sheet_data.nrows):

            row_data_list = sheet_data.row_values(row_index)[:40]

            if len(row_data_list) < 27:
                continue
            elif not row_data_list[rkrq_index]:
                continue
            elif row_data_list[wlzt_index]:
                continue
            else:
                for index,cell in enumerate(row_data_list):
                    if sheet_tbody_list[index] in self.table_name_list:
                        col = self.table_name_list.index(sheet_tbody_list[index])
                        # 日期格式转换
                        if col in [10,11,21,22,26,34,35,36,38] and isinstance(cell,(int,float)) and float(cell) > 0.0:
                            cell = xlrd.xldate_as_tuple(cell, 0)
                            self.ws.write(self.row_indext, col, datetime.datetime(cell[0],cell[1],cell[2]),self.dateFormat)
                        else:
                            self.ws.write(self.row_indext, col, cell)

                self.row_indext += 1
                self.old_content.save(f"工作簿{self.tody}.xlsx")

    def verify_tbody(self,sheet_tbody_list,sheet_name):
        # 校验表头
        for cell in self.table_name_list[:-1]:
            if cell not in sheet_tbody_list[:27]:
                error_msg = f"文件：【{self.xlsx_pwd_split}】，sheet:【{sheet_name}】模板表头和待整理表表头不匹配。不存在【{cell}】列,请查看修改"
                raise Exception(error_msg)

    def deal_info(self):
        with open(f"写入失败文件{self.tody}.txt","w") as f:
            error_file_num = len(set([sheet[0] for sheet in self.error_sheet_list]))
            f.writelines(f"失败文件数共计：【{error_file_num}】\n失败sheet数共计：【{len(self.error_sheet_list)}】\n空白sheet数共计：【{len(self.empty_sheet_list)}】\n详细情况如下：\n")
            f.writelines(f"以下是写入失败的：\n"+"-"*100+"\n")
            for row_data in self.error_sheet_list:
                f.writelines("\n".join(row_data)+"\n"+"&"*100+"\n")
            f.writelines(f"以下是空表：\n" + "-" * 100+"\n")
            for row_data in self.empty_sheet_list:
                f.writelines("\n".join(row_data)+"\n")

main()
