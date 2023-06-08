import os,re,xlrd,time
from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd
import numpy as np
import shutil
from sklearn.feature_extraction.text import TfidfVectorizer
import win32com.client as win32

import warnings
warnings.filterwarnings('ignore')





# 时间格式转化
def T_date(df):
    for i in df.columns:
        if df[i].dtype == 'datetime64[ns]':
            df[i] = df[i].apply(lambda x: str(pd.to_datetime(x).date()))

def check_key(data_path):
    print('正在检查 关键字段是否缺失...')
    print('='*200)
    msg=1
    new_data_path = data_path + '处理成功'
    list1=['销售日期','客户名称','产品名称','规格','产品单位','数量','生产厂家']
    for dirpath,dirname,filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8]=='SAL':
                file_path=os.path.join(dirpath,fname)
                df=pd.read_excel(file_path,dtype='object')
                for i in list1:
                    if i not in list(df.columns):
                        msg+=1
                        print(f'{i} 字段缺失->| {file_path}')
                        print('-'*200)
    return msg




#去除非必填字段，选填字段
def reduce_data(data_path):
    new_data_path=data_path+'处理成功'
    df_final=pd.DataFrame()
    for dirpath,dirname,filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path=os.path.join(dirpath,fname)
            if fname[5:8]=='SAL':
                # print(file_path)
                df=pd.read_excel(file_path,dtype='object')
                list1=['销售日期','客户名称','产品名称','规格','产品单位','数量','批号','单价','金额','送货地址','生产厂家','产品编码','客户编码']
                list2=list(df.columns)
                column_names=[column_name for column_name in list1 if column_name in list2]
                df_final=df[column_names]
                df.dropna(how='any',axis=0,inplace=True)
                # df.to_excel(file_path, index=False)
                df_final=df_final[(df_final['数量']!=0)&(df_final['数量']!='0')&(df_final['数量']!='&nbsp;')\
                                  &(df_final['数量'].notnull())&(df_final['数量']!='合计：')]
                T_date(df_final)
                df_final.to_excel(file_path,index=False)



def del_flows(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'SAL':
                # print(file_path)
                df = pd.read_excel(file_path, dtype='object')
                df = df[(df['销售日期'].notnull()) & (df['销售日期'] != '合计') & (df['销售日期'] != '合计：') & (df['销售日期'] != 'NaT')&(df['销售日期'] != '业务日期')]
                df = df[(df['产品单位'].notnull()) & (df['产品单位'] != '----------')]
                df = df[(df['客户名称'].notnull()) & (df['客户名称'] != '~')]
                df.to_excel(file_path, index=False)


def sale_date_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            # print(file_path)
            df = pd.read_excel(file_path, dtype='object')
            T_date(df)
            if fname[5:8] == 'SAL' and df.shape[0] > 0:
                list3 = []
                for i in df['销售日期']:
                    try:
                        i = pd.to_datetime(str(i)[:10]).strftime('%Y-%m-%d')
                        list3.append(i)
                    except:
                        try:
                            # print(file_path)
                            i = pd.to_datetime(str(i)[:9]).strftime('%Y-%m-%d')
                            list3.append(i)
                        except:
                            floder_name = data_path + '处理失败' + '\\' + '销售日期' + '\\' + '\\'.join(
                                file_path.split('\\')[len(data_path.split('\\')):-1])
                            if not os.path.exists(floder_name):
                                os.makedirs(floder_name)
                            new_file_path = floder_name + '\\' + fname
                            try:
                                shutil.move(file_path, new_file_path)
                                print(f'销售日期 清洗失败->|{new_file_path}')
                                print('-' * 200)
                            except:
                                pass
                    try:
                        df['销售日期'] = list3
                        df.to_excel(file_path, index=False)
                    except:
                        pass
                    # try:
                    #     df['销售日期'] = list3
                    #     df.to_excel(file_path, index=False)
                    # except:
                    #     floder_name = data_path + '处理失败' + '\\' + '销售日期' + '\\' + '\\'.join(
                    #         file_path.split('\\')[len(data_path.split('\\')):-1])
                    #     if not os.path.exists(floder_name):
                    #         os.makedirs(floder_name)
                    #     new_file_path = floder_name + '\\' + fname
                    #     try:
                    #         shutil.move(file_path, new_file_path)
                    #         print(f'销售日期 清洗失败->|{new_file_path}')
                    #         print('-' * 200)
                    #     except:
                    #         pass


def check_data(data_path):
    new_data_path = data_path + '处理成功'
    concat_df = pd.DataFrame()
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'SAL':
                df = pd.read_excel(file_path, dtype='object')
                #         print(file_path)
                #         pattern=re.compile('[\u4e00-\u9fa5].*[司站队部房店肃心行院]')
                #         result=pattern.findall(fname)[0]

                # list1 = ['销售日期', '客户名称', '产品名称', '规格', '产品单位', '数量', '批号', '单价', '金额', '送货地址', '生产厂家']
                # list2 = list(df.columns)
                # column_names = [column_name for column_name in list1 if column_name in list2]
                # df_final = df[column_names]
                df['经销商'] = file_path.split('\\')[-2].split('_')[1]
                concat_df = pd.concat([concat_df, df], axis=0)
    concat_df = concat_df[['销售日期', '客户编码','客户名称', '产品编码','产品名称', '规格', '产品单位', '数量', '单价', '金额', '批号', '生产厂家', '送货地址', '经销商']]
    concat_df.to_excel(data_path+'数据合并SAL.xlsx', index=False)


def add_factory_manue(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'SAL':
                df = pd.read_excel(file_path, dtype='object')
                df['生产厂家']=df['生产厂家'].fillna('xy')
                df.to_excel(file_path,index=False)

if __name__ == '__main__':
    data_path = r'C:\Users\guodingyu\Desktop\工具\SHXY_CLEAN\接收文件二级商\20220922'
    time_start = time.time()  # 记录开始时间
    msg=check_key(data_path) #检查必填字段是否缺失
    if msg==1:
        print('关键字段检查完成，开始清洗数据...')
        reduce_data(data_path)
        del_flows(data_path)
        sale_date_clean(data_path)
        add_factory_manue(data_path) #生产厂家缺失的补 xy
        check_data(data_path)
    else:
        print('关键字段缺失，请检查...')
    time_end = time.time()
    time_sum = time_end - time_start
    print(f'数据清洗完成，程序运行->| {round(time_sum, 2)}s ')

