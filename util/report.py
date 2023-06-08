import pandas as pd
import os, re
import datetime
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
import itertools


class ReportCleanData(object):
    """
    :param data_path:遍历文件夹下含‘接收文件’的文件夹，生成分拣报告
    """

    def __init__(self, data_path):
        clean_path = r'.\data\xy模糊识别库.xlsx'
        self.clean_path = clean_path
        self.data_path = data_path

    def makereport(self):
        data_clean = pd.read_excel(self.clean_path, usecols=['经销商文件名', '经销商编码', '经销商名称', 'DPS名称'])
        list1 = []
        list2 = []
        for floder in os.listdir(self.data_path):
            if floder.__contains__('接收文件'):
                report_data = os.path.join(self.data_path, floder)
                for dirpath, dirname, filenames in os.walk(report_data):
                    for fname in filenames:
                        file_path = os.path.join(dirpath, fname)
                        pattern = re.compile('[\u4e00-\u9fa5].*[司站队部房店肃心行夏院药]')
                        result = pattern.findall(file_path.split(sep='\\')[-2])
                        if len(result) == 0:
                            # print(f'清洗失败->| {file_path}')
                            break
                        else:
                            list1 = [file_path.split(sep='\\')[-4], '\\'.join(file_path.split(sep='\\')[-3:-1]),
                                     file_path.split(sep='\\')[-1], result[0]]
                            list2.append(list1)
        df_final = pd.DataFrame(list2, columns=['日期', '文件夹名称', '文件名', '经销商文件名'])
        df_final = pd.merge(df_final, data_clean, how='left', on='经销商文件名')
        df_final['经销商编码'] = df_final['经销商编码'].fillna('未识别')
        df_final['邮件主题'] = ''
        df_final['是否有效'] = 'Y'
        df_final['备注'] = ''
        df_final = df_final[
            ['日期', '邮件主题', '文件夹名称', '文件名', '经销商编码', '经销商名称', '是否有效', '备注',
             '经销商文件名', 'DPS名称']]
        df_final['日期'] = df_final['日期'].map(lambda x: str(pd.to_datetime(x.replace('接收文件', '')).date()))
        report_month = str(datetime.datetime.now().month) + '月'  # 获取当前月份
        report_date = ''.join(str(datetime.datetime.now())[:10].split(sep='-'))  # 获取分拣报告时间，默认为当天时间

        report_path = self.data_path + '\\' + report_month + self.data_path.split(sep='\\')[-1] + report_date + '.xlsx'
        # df_final.to_excel(report_path, index=False)
        # 定义EXCEL样式
        writer = pd.ExcelWriter(report_path, engine='openpyxl')
        df_final.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        font = Font(name='微软雅黑', bold=True, color='f7f7f7')
        alignment = Alignment(vertical='top', wrap_text=True)
        pattern_fill = PatternFill(fill_type='solid', fgColor='00b0f0')
        side = Side(style='thin')
        border = Border(left=side, right=side, top=side, bottom=side)
        for cell in itertools.chain(*worksheet['A1:J1']):
            cell.font = font
            cell.alignment = alignment
            cell.fill = pattern_fill
            cell.border = border
        worksheet.column_dimensions['A'].width = 12
        worksheet.column_dimensions['B'].width = 10
        worksheet.column_dimensions['C'].width = 50
        worksheet.column_dimensions['D'].width = 35
        worksheet.column_dimensions['E'].width = 15
        worksheet.column_dimensions['F'].width = 30
        worksheet.column_dimensions['G'].width = 10
        worksheet.column_dimensions['H'].width = 10
        worksheet.column_dimensions['I'].width = 30
        worksheet.column_dimensions['J'].width = 40
        writer.save()
        writer.close()
        return df_final

# report_data_path=r'D:\shxy\接收文件一级商'
# report_obj=ReportCleanData(data_path=report_data_path)
# report=report_obj.makereport()
