import pandas as pd
import shutil
import calendar
import datetime
import re,os
import warnings
warnings.filterwarnings('ignore')

def T_date(df):
    for i in df.columns:
        if df[i].dtype == 'datetime64[ns]':
            df[i] = df[i].apply(lambda x: str(pd.to_datetime(x).date()))


#云南省医药有限公司
def shxy_yunnanshengyiyao(file_path):
    df=pd.read_excel(file_path,dtype='object')
    df=df[df['品名规格'].notnull()]#删除空行
    #提取产品名称
    df['产品']=df['产品'].loc[df[df['品名规格']=='开票日期'].index-1]=df['品名规格'].loc[df[df['品名规格']=='开票日期'].index-1]
    df['产品']=df['产品'].fillna(method='ffill')
    #提取规格
    df['规格']=df['产品'].str.extract('(\d.*)')
    #提取产品单位
    df['产品单位']=df['单位'].fillna(method='ffill')
    #提取销售数量
    df['数量']=df['Unnamed: 2']
    #提取生产厂家
    df['生产厂家']=df['厂商'].fillna(method='ffill')
    #提取客户名称
    df['客户名称']=df['Unnamed: 1']
    #提取销售日期
    df['销售日期']=df['品名规格']
    #提取库存数量
    df['库存数量']=df['库存']
    #提取销售数据
    sal=df.iloc[:,9:16]
    #提取库存数据
    inv=df[['产品','规格','产品单位','生产厂家','库存数量']]
    #销售数据清洗
    sal=sal[(sal['数量'].notnull())&(sal['数量']!='销售数量')&(sal['客户名称'].notnull())]
    sal['销售日期']=sal['销售日期'].map(lambda x: str(x)[:10])
    sal['产品名称']=sal['产品'].str.split(' ',expand=True)[0]
    sal.drop('产品',axis=1,inplace=True)
    sal=sal[['销售日期','客户名称','产品名称','规格','数量','产品单位','生产厂家']]
    #库存数据清洗
    inv=inv[(inv['库存数量'].notnull())&(inv['库存数量']!='库存')]
    inv['产品名称']=inv['产品'].str.split(' ',expand=True)[0]
    inv.drop('产品',axis=1,inplace=True)
    inv=inv[['产品名称','规格','产品单位','生产厂家','库存数量']]
    #构造库存文件路径
    inv_name=file_path.split(sep='\\')[-1]
    pattern=re.compile('\d{8}')
    result=pattern.findall(inv_name)
    inv_fname=f'SHXY_INV_MON_{result[0]}ZC1.xlsx'
    sal_fname=f'SHXY_SAL_MON_{result[0]}ZC1.xlsx'
    inv_path='\\'.join(file_path.split(sep='\\')[:-1])+'\\'+inv_fname
    sal_path='\\'.join(file_path.split(sep='\\')[:-1])+'\\'+sal_fname
    # 删除原文件
    os.remove(file_path)
    #写入文件
    T_date(sal)
    T_date(inv)
    sal.to_excel(sal_path,index=False)
    inv.to_excel(inv_path,index=False)

#广西二级经销商处理
def shxy_guangxi_clean(new_data_path):
    # new_data_path=data_path+'处理成功'
    # data_path=data_path+'处理结果'
    dicta={
    '000200081333_柳州桂中大药房连锁有限责任公司_SHXY':'柳州桂中大药房连锁有限责任公司'
    ,'005259123338_广西桂林柳药药业有限公司_SHXY':'桂林柳药药业有限公司'
    ,'005259123370_广西玉林柳药药业有限公司_SHXY':'玉林柳药药业有限公司'
    ,'005259123484_南宁柳药药业有限公司_SHXY':'南宁柳药药业有限公司'
    ,'005259125645_广西梧州柳药药业有限公司_SHXY':'梧州柳药药业有限公司'
    ,'005263540279_百色柳药药业有限公司_SHXY':'百色柳药药业有限公司'
    ,'005263540280_贵港柳药药业有限公司_SHXY':'贵港柳药药业有限公司'
    ,'005264798543_河池市柳药药业有限公司_SHXY':'河池柳药药业有限公司'
    ,'005336720933_广西贺州柳药药业有限公司_SHXY':'贺州柳药药业有限公司'
    ,'005380297845_广西来宾柳药药业有限公司_SHXY':'来宾柳药药业有限公司'
    ,'005387814610_广西钦州柳药药业有限公司_SHXY':'钦州柳药药业有限公司'
    ,'005406478219_广西桂林柳药弘德医药有限公司_SHXY':'桂林柳药弘德医药'
    ,'005416079999_广西北海柳药药业有限公司_SHXY':'北海柳药药业有限公司'
    ,'005418981809_广西金夫康医药有限公司_SHXY':'广西金夫康医药有限公司'
    ,'005450809201_广西崇左柳药药业有限公司_SHXY':'崇左柳药药业有限公司'
    }
    for dirpath,dirname,filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path=os.path.join(dirpath,fname)
            floder_name=file_path.split(sep='\\')[-2]#文件夹名称
            if new_data_path.split(sep='\\')[-2]=='接收文件二级商':
                for i in dicta:
                    if floder_name==i:
                        df=pd.read_excel(file_path,dtype='object')
                        for header_name in df.columns:
                            if header_name=='销售部门':
                                df_final=df[df['销售部门']==dicta[i]]
                                try:
                                    df_final.drop(['商品编码','属性','卡号','卡类别'],axis=1,inplace=True)
                                except:
                                    df_final.rename(columns={'通用名':'产品名称','单位':'产品单位','制单日期':'销售日期'},inplace=True)
                                    df_final.to_excel(file_path,index=False)
    # print('广西二级商清洗完成...')
    # print('='*200)

#浙江金益医药有限公司
def zhejiangjinyi(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.columns = ['销售日期', '未知1', '未知2', '品名', '规格', '未知3', '产品单位', '数量', '批号', '单价', '金额', '客户名称']
    df = df[['销售日期', '品名', '规格', '产品单位', '数量', '批号', '单价', '金额', '客户名称']]
    df.rename(columns={'品名': '产品名称'}, inplace=True)
    df = df[['销售日期', '客户名称', '产品名称', '规格', '产品单位', '批号', '数量', '单价', '金额']]
    T_date(df)
    df.to_excel(file_path, index=False)

# 四川雅安金海堂药业有限责任公司
def yananjinhaitang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    date_time = file_path.split('\\')[-1].split('_')[3][:8]

    num_sal = df[df.iloc[:, 0] == '商品销售流向'].index[0]
    num_inv = df[df.iloc[:, 0] == '库存'].index[0]

    df_pur = df.iloc[:num_sal, :]
    df_sal = df.iloc[num_sal:num_inv, :].reset_index(drop=True)
    df_inv = df.iloc[num_inv:, :].reset_index(drop=True)
    df_pur.columns = df_pur.iloc[2, :].reset_index(drop=True)
    df_pur.dropna(how='any', inplace=True)
    df_pur = df_pur[df_pur['单据日期'] != '单据日期']
    df_pur.reset_index(drop=True, inplace=True)
    df_pur.rename(columns={'单据日期': '销售日期', '商品名称': '产品名称', '商品规格': '规格'
        , '单位': '产品单位'}, inplace=True)
    df_sal.columns = df_sal.iloc[2, :].reset_index(drop=True)
    df_sal.dropna(how='any', inplace=True)
    df_sal = df_sal[df_sal['单据日期'] != '单据日期']
    df_sal.reset_index(drop=True, inplace=True)
    df_sal.rename(columns={'单据日期': '销售日期', '往来单位': '客户名称', '商品名称': '产品名称', '商品规格': '规格'
        , '单位': '产品单位'}, inplace=True)
    df_inv.columns = df_inv.iloc[1, :].reset_index(drop=True)
    df_inv.dropna(how='all', inplace=True, axis=0)
    df_inv.dropna(how='all', inplace=True, axis=1)
    df_inv = df_inv[(df_inv['商品名称'] != '商品名称') & (df_inv['商品名称'] != '库存')]
    df_inv.rename(columns={'商品名称': '产品名称', '商品规格': '规格', '单位': '产品单位', '库存数量': '数量'}, inplace=True)
    df_sal['销售日期'] = df_sal['销售日期'].map(lambda x: str(x)[:10])
    df_pur['销售日期'] = df_pur['销售日期'].map(lambda x: str(x)[:10])
    df_inv['库存日期'] = pd.to_datetime(date_time)
    df_inv['库存日期'] = df_inv['库存日期'].map(lambda x: str(x)[:10])
    # 构造文件路径
    pur_path = '\\'.join(file_path.split('\\')[:-1]) + '\\' + 'SHXY_PUR_MON_' + date_time + 'ZC2' + '.xlsx'
    inv_path = '\\'.join(file_path.split('\\')[:-1]) + '\\' + 'SHXY_INV_MON_' + date_time + 'ZC2' + '.xlsx'
    sal_path = '\\'.join(file_path.split('\\')[:-1]) + '\\' + 'SHXY_SAL_MON_' + date_time + 'ZC2' + '.xlsx'
    df_pur.to_excel(pur_path, index=False)
    df_inv.to_excel(inv_path, index=False)
    df_sal.to_excel(sal_path, index=False)
    os.remove(file_path)


#华东医药股份有限公司中成药分公司: 多流向拆分
def zhongchengyao(file_path):
    df = pd.read_excel(file_path,dtype='object')
    df.rename(columns={'供应数量':'数量'},inplace=True)
    sal_df=df[(df[' 单据类型']!='进货')&(df[' 单据类型']!='进退')]
    pur_df=df[(df[' 单据类型']=='进货')|(df[' 单据类型']=='进退')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>1:
        T_date(pur_df)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.to_excel(file_path,index=False)


#浙江省嘉信医药股份有限公司: 多流向拆分
def zhejiangjiaxin(file_path):
    df = pd.read_excel(file_path,dtype='object')
    sal_df=df[(df['单据类型']!='进货')&(df['单据类型']!='进退')]
    pur_df=df[(df['单据类型']=='进货')|(df['单据类型']=='进退')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.to_excel(file_path,index=False)

#济南市爱新卓尔医药有限责任公司: 多流向拆分
def aixinzhuoer(file_path):
    df = pd.read_excel(file_path,dtype='object')
    sal_df=df[(df['摘要']!='采购入库单')]
    pur_df=df[(df['摘要']=='采购入库单')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.to_excel(file_path,index=False)

#宣城市宣州区昭亭路医药有限公司: 多流向拆分
def zhaotinglu(file_path):
    df = pd.read_excel(file_path,dtype='object')
    if '摘要' in list(df.columns):
        sal_df=df[(df['摘要']!='采购入库')]
        pur_df=df[(df['摘要']=='采购入库')]
        floder_name='\\'.join(file_path.split('\\')[:-1])
        fname=file_path.split('\\')[-1]
        num=fname.split('_')[3]
        pur_fname='SHXY_PUR_MON_'+num
        pur_path=floder_name+'\\'+pur_fname
        if pur_df.shape[0]>=1:
            T_date(pur_df)
            pur_df.to_excel(pur_path,index=False)
        T_date(sal_df)
        sal_df.to_excel(file_path,index=False)
    else:
        pass

#福建鑫天健医药有限公司: 多流向拆分
def xintianjian(file_path):
    try:
        df = pd.read_excel(file_path,dtype='object')
        sal_df=df[(df['摘要']!='采购入库单')]
        pur_df=df[(df['摘要']=='采购入库单')]
        floder_name='\\'.join(file_path.split('\\')[:-1])
        fname=file_path.split('\\')[-1]
        num=fname.split('_')[3]
        pur_fname='SHXY_PUR_MON_'+num
        pur_path=floder_name+'\\'+pur_fname
        if pur_df.shape[0]>=1:
            T_date(pur_df)
            pur_df.to_excel(pur_path,index=False)
        T_date(sal_df)
        sal_df.to_excel(file_path,index=False)
    except:
        df = pd.read_excel(file_path, dtype='object')
        df.rename(columns={'单位名称':'客户名称','日期':'销售日期','单位':'产品单位'},inplace=True)
        df.to_excel(file_path,index=False)


#嘉兴市英特医药有限公司: 多流向拆分
def jiaxingyingte(file_path):
    df = pd.read_excel(file_path,dtype='object')
    sal_df=df[(df['单据类型']!='进货')]
    pur_df=df[(df['单据类型']=='进货')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.rename(columns={'产地':'生产厂家'},inplace=True)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.rename(columns={'产地': '生产厂家'},inplace=True)
    sal_df.to_excel(file_path,index=False)
# def jiaxingyingte_pur(file_path):
#     df = pd.read_excel(file_path, dtype='object').rename(columns={'产地':'生产厂家'})
#     T_date(df)
#     df.to_excel(file_path, index=False)

#浙江省英特药业有限责任公司杭州新特药分公司: 多流向拆分
def hangzhouxinte(file_path):
    df = pd.read_excel(file_path,dtype='object')
    sal_df=df[(df['单据类型']!='进货')&(df['单据类型']!='进退')]
    pur_df=df[(df['单据类型']=='进货')|(df['单据类型']=='进退')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.to_excel(file_path,index=False)

#华东医药台州有限公司: 多流向拆分
def huadongtaizhaou(file_path):
    df = pd.read_excel(file_path,dtype='object')
    sal_df=df[(df[' 单据类型']!='进货')&(df[' 单据类型']!='进退')]
    pur_df=df[(df[' 单据类型']=='进货')|(df[' 单据类型']=='进退')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.to_excel(file_path,index=False)

#郴州凯程医药有限公司: 多流向拆分
def kaichen(file_path):
    df = pd.read_excel(file_path,dtype='object')
    sal_df=df[(df['单据类型']!='验收入库单')]
    pur_df=df[(df['单据类型']=='验收入库单')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.to_excel(file_path,index=False)

#浙江省大德药业集团浙江省医药有限公司: 多流向拆分
def zhejiangdade(file_path):
    df = pd.read_excel(file_path,dtype='object')
    sal_df=df[(df['类型']!='采购入库')]
    pur_df=df[(df['类型']=='采购入库')]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.to_excel(file_path,index=False)
#金华市东阳医药药材有限公司: 多流向拆分
def jinhuadongyang(file_path):
    try:
        df = pd.read_excel(file_path, dtype='object')
        sal_df=df[(df['注译']!='药品入库单')]
        pur_df=df[(df['注译']=='药品入库单')]
        floder_name='\\'.join(file_path.split('\\')[:-1])
        fname=file_path.split('\\')[-1]
        num=fname.split('_')[3]
        pur_fname='SHXY_PUR_MON_'+num
        pur_path=floder_name+'\\'+pur_fname
        if pur_df.shape[0]>=1:
            T_date(pur_df)
            pur_df.rename(columns={'产品':'产品名称'},inplace=True)
            pur_df.to_excel(pur_path,index=False)
        T_date(sal_df)
        sal_df.to_excel(file_path,index=False)
    #2022年10月9日更新
    except:
        df = pd.read_excel(file_path, dtype='object',header=1)
        df.rename(columns={'购入客户名称':'客户名称'},inplace=True)
        df.to_excel(file_path,index=False)


#国药控股海南省鸿益有限公司：销售日期为制单日期
def hainanhongyi(file_path):
    df = pd.read_excel(file_path,dtype='object')
    df.rename(columns={'制单日期':'销售日期'},inplace=True)
    df.drop('出库日期',axis=1,inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#福建省中源医药有限公司：销售流向：客户名称不易区分
def fujianzhongyuan(file_path):
    df = pd.read_excel(file_path,dtype='object')
    df.rename(columns={'客商ID':'客户编码','客商':'客户名称'},inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)
#四川省本草堂药业有限公司
def sichuanbencao(file_path):
    df = pd.read_excel(file_path,dtype='object')
    df.rename(columns={'客户ID':'客户编码','客户':'客户名称','货品ID':'产品编码'},inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)
#陕西省天士力医药有限公司
def shanxitianlishi(file_path):
    df = pd.read_excel(file_path,dtype='object')
    df.rename(columns={'客户ID':'客户编码','客户':'客户名称','货品ID':'产品编码'},inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)
#江苏宏康医药有限责任公司
def jiangsuhongkang(file_path):
    df = pd.read_excel(file_path,dtype='object')
    df=df.iloc[:,:11]
    list1=['销售日期','单据编号','客户名称','产品编码','产品名称','规格','生产厂家','采购数量','数量','批号','有效期']
    df.columns=list1
    df['采购数量']=df['采购数量'].map(lambda x: float(x))
    sal_df=df[df['采购数量']==0]
    pur_df=df[df['采购数量']>0]
    floder_name='\\'.join(file_path.split('\\')[:-1])
    fname=file_path.split('\\')[-1]
    num=fname.split('_')[3]
    pur_fname='SHXY_PUR_MON_'+num
    pur_path=floder_name+'\\'+pur_fname
    if pur_df.shape[0]>=1:
        T_date(pur_df)
        pur_df.drop('数量',axis=1,inplace=True)
        pur_df.rename(columns={'销售日期':'采购日期'},inplace=True)
        pur_df.to_excel(pur_path,index=False)
    T_date(sal_df)
    sal_df.drop('采购数量',axis=1,inplace=True)
    sal_df.to_excel(file_path,index=False)

#山东康诺盛世医药有限公司
def kangnuoshengshi(file_path):
    df = pd.read_excel(file_path,dtype='object')
    df.rename(columns={'货品ID':'产品编码','货品名称':'产品名称','货品规格':'规格','销售数量(汇总)':'数量'},inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#江西省五洲医药营销有限公司
def jiangxiwuzhou(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['销售日期'] = df['制单日'].map(lambda x: str('20' + x))
    df['效期'] = df['有效期至'].map(lambda x: str('20' + x))
    df.rename(columns={'客户': '客户名称'}, inplace=True)
    df.drop(['制单日', '有效期至'], axis=1, inplace=True)
    # list1 = ['制单日期', '客户名称', '产品名称', '规格', '生产厂家', '产品单位', '批号', '有效期', '数量', '生产日期', '国药准字', '销售日期']
    # df.columns = list1
    # df = df[['销售日期', '客户名称', '产品名称', '规格', '产品单位', '数量', '批号', '生产厂家']]
    # T_date(df)
    df.to_excel(file_path, index=False)

#广西华泰药业有限公司
def guangxihuatai(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['药品信息'].fillna(method='ffill', inplace=True)
    df['规格'] = df['药品信息']
    df['生产厂家'] = df['药品信息']
    df['产品单位'] = '缺失'
    df.rename(columns={'药品信息': '产品名称', '制单日期': '销售日期', '客商名称': '客户名称', '客商号': '客户代码',
                       '销售数量': '数量', '库存': '库存数量'}, inplace=True)
    df = df[df['客户名称'] != '期初结余库存']
    df_sal = df[['销售日期', '客户名称', '产品名称', '规格', '产品单位', '数量', '批号', '生产厂家']]
    df_inv = df[['产品名称', '规格', '产品单位', '库存数量', '批号', '生产厂家']]

    floder_name = '\\'.join(file_path.split('\\')[:-1])
    fname = file_path.split('\\')[-1]
    num = fname.split('_')[3]
    pur_fname = 'SHXY_PUR_MON_' + num
    pur_path = floder_name + '\\' + pur_fname
    df_inv['库存日期'] = num[:8]
    df_sal.to_excel(file_path, index=False)
    df_inv.to_excel(pur_path, index=False)


#江苏吴中医药销售有限公司_
def jiangsuwuzhong(file_path):
    df=pd.read_excel(file_path,dtype='object',header=1)
    df['销售日期']=df['销售日期'].map(lambda x: '20'+x)
    df.columns=['销售日期','客户代码','客户名称','产品编码','产品名称','规格','生产厂家','产品单位',
                '批号','有效期','数量','单价','金额']
    df=df[['销售日期', '客户名称', '产品名称', '规格', '产品单位', '数量', '批号', '生产厂家','客户代码','产品编码']]
    T_date(df)
    df.to_excel(file_path, index=False)

#重庆市长圣医药有限公司
def chongqingchangsheng(file_path):
    df=pd.read_excel(file_path,dtype='object',header=1)
    df.rename(columns={'单位':'客户名称'},inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)
#湖北人福医药集团有限公司
def hubeirenfu(file_path):
    df=pd.read_excel(file_path,dtype='object')
    df.rename(columns={'销售商': '客户名称'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#江西省南华医药有限公司
def jiangxihuanan(file_path):
    df=pd.read_excel(file_path,dtype='object')
    df.rename(columns={'出库日期 ': '销售日期'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)
#上药科园信海医药湖北有限公司
def keyuanxinhai(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.rename(columns={'已出发票数量': '数量'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#云南医药工业销售有限公司
def yunnanyiyaogongye(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.rename(columns={'实际出库数量': '数量','单位':'产品单位'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#武威神洲医药有限责任公司
def wuweishenzhou(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.rename(columns={'单位全名': '供应商名称'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#四川腾龙医药有限公司
def sichuantenglong(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.rename(columns={'单位名称': '客户名称'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#北京市科园信海医药经营有限公司
def beijingkeyuan(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.rename(columns={'已出发票数量': '数量'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#厦门片仔癀宏仁医药有限公司
def pianzaihaunghongren(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.rename(columns={'流向数量': '数量'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#国药集团西南医药有限公司
def guoyaojituanxinan(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df.rename(columns={'开单日期(含时分秒)': '销售日期'}, inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#江西省广力药业有限公司
def jiangxiguangli(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['销售日期']=df['制单日'].map(lambda x: str('20'+x))
    T_date(df)
    df.to_excel(file_path, index=False)

#江西吉安医药有限公司
def jiangxijian(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['销售日期']=df['制单日'].map(lambda x: str('20'+x))
    T_date(df)
    df.to_excel(file_path, index=False)

#山西省康美徕医药有限公司
def shanxikangmeilai(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['生产厂家']=df['厂牌']
    # df['厂牌']=df['生产厂家'].map(lambda x: str('20'+x))
    T_date(df)
    df.to_excel(file_path, index=False)

#浙江省恩泽医药有限公司
def zhejiangenze(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['生产厂家'] = df['VCMANUFACTURER']
    T_date(df)
    df.to_excel(file_path, index=False)

#山东省容大医药有限公司
def shandongronda(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['生产厂家'] = df['生产企业']
    T_date(df)
    df.to_excel(file_path, index=False)

#河北润祥医药有限责任公司
def hebeirunxiang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    if '数量' not in list(df.columns):
        df['数量'] = df['结算数量']
    T_date(df)
    df.to_excel(file_path, index=False)

#菏泽海王医药有限公司
def hezehaiwang(file_path):
    df = pd.read_excel(file_path, dtype='object',header=1)
    df['生产厂家'] = df['生产企业名称']
    T_date(df)
    df.to_excel(file_path, index=False)

#兰州市强生医药有限责任公司
def lanzhouqiangsheng(file_path):
    fname=file_path.split('\\')[-1]
    if fname[5:8]=='SAL':
        df = pd.read_excel(file_path, dtype='object')
        df['销售日期'] = df.iloc[:,2].map(lambda x: x.replace(' ', ''))
        df['效期'] = df['失效日期'].map(lambda x: x.replace(' ', ''))
        df.drop('日期', axis=1, inplace=True)
        df.drop('日期.1', axis=1, inplace=True)
        T_date(df)
        df.to_excel(file_path, index=False)
    elif fname[5:8]=='PUR':
        df = pd.read_excel(file_path, dtype='object')
        df['采购日期'] = df['日期'].map(lambda x: x.replace(' ', ''))
        df['效期'] = df['效期'].map(lambda x: x.replace(' ', ''))
        # df.drop('日期', axis=1, inplace=True)
        T_date(df)
        df.to_excel(file_path, index=False)

#梅州市卫发医药有限公司
def meizhouweifa(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['采购日期'] = df['制单日期']
    df.rename(columns={'供应商名':'供应商名称'},inplace=True)
    T_date(df)
    df.to_excel(file_path, index=False)

#南宁华御堂医药公司
def nanninghuayutang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['入库数量']
    T_date(df)
    df.to_excel(file_path, index=False)

#安徽东升医药物流有限公司
def anhuidongshen(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['库房结存数量']
    T_date(df)
    df.to_excel(file_path, index=False)

#湖南博瑞药业有限公司
def hunanborui(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['可用数量']
    T_date(df)
    df.to_excel(file_path, index=False)
#广西桂林汇通药业有限公司
def guilinhuitong(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['入库数量']
    T_date(df)
    df.to_excel(file_path, index=False)
#浙江广为医药有限公司
def zhejiangguangwei(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['产品名称'] = df['通用名(商品名)\规 格']
    df['规格'] = df['通用名(商品名)\规 格']
    T_date(df)
    df.to_excel(file_path, index=False)
#邢台万邦医药有限责任公司
def xingtaiwanbang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    T_date(df)
    df.to_excel(file_path, index=False)
def xingtaiwanbang_pur(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['采购日期'] = df['记账日期']
    T_date(df)
    df.to_excel(file_path, index=False)

#安徽省亳州市医药供销有限公司
def anhuihaozhou(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['下游客户名称']
    T_date(df)
    df.to_excel(file_path, index=False)
#华润芜湖医药有限公司
def huarunwuhu(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#安徽省华源医药股份有限公司
def anhuihuayuan(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['主数量']
    T_date(df)
    df.to_excel(file_path, index=False)
#上药控股江苏股份有限公司
def shangyaojiangsu(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['订单数量']
    T_date(df)
    df.to_excel(file_path, index=False)

#国药控股河南上蔡分公司
def henanshangcai(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['产品名称'] = df['品名规格']
    T_date(df)
    df.to_excel(file_path, index=False)
#广汉市吉昌药业有限公司
def guanghanjichang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['相关单位名称']
    T_date(df)
    df.to_excel(file_path, index=False)
#东辽县医药药材有限责任公司
def dongliaoyiyao(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['购货客户']
    T_date(df)
    df.to_excel(file_path, index=False)
#淮滨县天一药品经营有限公司
def huaibintianyi(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位名称']
    T_date(df)
    df.to_excel(file_path, index=False)
#陕西省恒庆医药有限公司
def shanxihengqing(file_path):
    df = pd.read_excel(file_path, dtype='object')
    try:
        df['客户名称'] = df['名称']
    except:
        df['客户名称'] = df['单位名称']
    T_date(df)
    df.to_excel(file_path, index=False)
#湖北省诚民天济药业有限公司
def hubeichengmin(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['对应门店']
    T_date(df)
    df.to_excel(file_path, index=False)
#黑龙江华辰大药房连锁有限公司
def heilongjianghuachen(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['机构']
    T_date(df)
    df.to_excel(file_path, index=False)
#人福医药钟祥有限公司
def renfuzhongxiang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['销售商']
    T_date(df)
    df.to_excel(file_path, index=False)
#华润周口医药有限责任公司
def huarunzhoukou(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#华润洛阳医药有限责任公司
def huarunluoyang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#华润三门峡医药有限责任公司
def huarunsanmenxia(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    df['客户代码'] = df['单位ID']
    T_date(df)
    df.to_excel(file_path, index=False)
#衡水市龙马医药贸易有限公司
def hengshuilongma(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#福建宏海药业有限公司
def fujianhonghai(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['供应商/客户']
    T_date(df)
    df.to_excel(file_path, index=False)
#河南九州通国华医药物流有限公司济源分公司
def guohuaqiyuan(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['购货单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#人福医药天门有限公司
def renfutianmen(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['销售商']
    T_date(df)
    df.to_excel(file_path, index=False)
#陕西铭川医药有限公司
def shanximingchuan(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['销售单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#三台县天诚医药有限公司
def santaitiancheng(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#无锡山禾集团健康参药连锁有限公司
def shanhejiankang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['相关单位名称']
    T_date(df)
    df.to_excel(file_path, index=False)
#重庆医药集团周口有限公司
def chongqingzhoukou(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['客户名称'] = df['单位']
    T_date(df)
    df.to_excel(file_path, index=False)
#舟山市存德医药有限公司
def zhoushancunde(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['供应数量']
    T_date(df)
    df.to_excel(file_path, index=False)
#陕西省医药孙思邈五星大药房连锁有限公司
def shanxisunsimiao(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['验收数量']
    T_date(df)
    df.to_excel(file_path, index=False)
#广西钦州市卫贸发展公司
def qinzhoushiweimao(file_path):
    try:
        df = pd.read_excel(file_path, dtype='object')
        df['数量'] = df['操作数量']
    except:
        df = pd.read_excel(file_path, dtype='object',header=1)
        df['数量'] = df['操作数量']
    T_date(df)
    df.to_excel(file_path, index=False)
#国药控股邓州有限公司
def guokongdengzhou(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['发生数量']
    T_date(df)
    df.to_excel(file_path, index=False)
#吉林省众鑫药业有限公司
def jilinzhongxin(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['汇总']
    T_date(df)
    df.to_excel(file_path, index=False)
#江苏致和堂医药物流有限公司
def jiangsuzhihetang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['数量'] = df['销售量']
    df['销售日期'] = df['批销日期']
    T_date(df)
    df.to_excel(file_path, index=False)
#威海市海王医药有限公司
def weihaihaiwang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['销售日期'] = df['签字日期']
    T_date(df)
    df.to_excel(file_path, index=False)
#河南省同和堂医药有限公司
def henantongrentang(file_path):
    df = pd.read_excel(file_path, dtype='object')
    df['销售日期'] = df['过账日期']
    T_date(df)
    df.to_excel(file_path, index=False)
#广西壮族自治区柳州市医药有限责任公司
def guangxiliuzhou(file_path):
    df = pd.read_excel(file_path, dtype='object', header=1)
    df.iloc[:, 0:4] = df.iloc[:, 0:4].fillna(method='ffill')
    df.drop(df[df['品名'].str.contains("小计")].index, inplace=True)
    df['生产厂家'] = df['生产厂家'].fillna("xy")
    T_date(df)
    df.to_excel(file_path, index=False)
def guangxiliuzhou_pur(file_path):
    df = pd.read_excel(file_path, dtype='object', header=1)
    df=df[(df['商品编码']!='[小计]')&(df['商品编码']!='合计')]
    # df.iloc[:, 0:4] = df.iloc[:, 0:4].fillna(method='ffill')
    # df.drop(df[df['品名'].str.contains("小计")].index, inplace=True)
    df['生产厂家'] = "xy"
    T_date(df)
    df.to_excel(file_path, index=False)



def make_floder(data_path,file_path,fname):
    new_floder_name = data_path + '处理失败' + '\\' + '特殊处理' + '\\' + '\\'.join(
        file_path.split('\\')[len(data_path.split('\\')):-1])
    if not os.path.exists(new_floder_name):
        os.makedirs(new_floder_name)
    new_file_path = new_floder_name + '\\' + fname
    shutil.move(file_path, new_file_path)




def shxy_spe(data_path):
    new_data_path=data_path+'处理成功'
    # time_start=time.time()
    # print('【正在处理特殊规则经销商，请稍后...】')
    # print('='*140)
    # data_path=data_path+'处理结果'
    # clean_path = r'C:\Users\guodingyu\Desktop\工具\SHXY_CLEAN\xy模糊识别库.xlsx'
    shxy_guangxi_clean(new_data_path)
    for dirpath,dirname,filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path=os.path.join(dirpath,fname)
            floder_name = file_path.split(sep='\\')[-2]
            fname=file_path.split('\\')[-1]
            pattern = re.compile('[\u4e00-\u9fa5].*[\u4e00-\u9fa5]')
            result = pattern.findall(floder_name)
            # print(file_path)
            # print(result[0])
            # print('-'*200)
            if result[0] == '云南省医药有限公司' and fname[5:8]=='SAL':
                try:
                    shxy_yunnanshengyiyao(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0]=='浙江金益医药有限公司' and fname[5:8]=='SAL':
                try:
                    zhejiangjinyi(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0]=='四川雅安金海堂药业有限责任公司' and fname[5:8]=='SAL':
                try:
                    yananjinhaitang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0]=='华东医药股份有限公司中成药分公司' and fname[5:8]=='SAL':
                try:
                    zhongchengyao(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '浙江省嘉信医药股份有限公司' and fname[5:8] == 'SAL':
                try:
                    zhejiangjiaxin(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '济南市爱新卓尔医药有限责任公司' and fname[5:8] == 'SAL':
                try:
                    aixinzhuoer(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '宣城市宣州区昭亭路医药有限公司' and fname[5:8] == 'SAL':
                try:
                    zhaotinglu(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '福建鑫天健医药有限公司' and fname[5:8] == 'SAL':
                try:
                    xintianjian(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '嘉兴市英特医药有限公司' and fname[5:8] == 'SAL':
                try:
                    jiaxingyingte(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '浙江省英特药业有限责任公司杭州新特药分公司' and fname[5:8] == 'SAL':
                try:
                    hangzhouxinte(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '华东医药台州有限公司' and fname[5:8] == 'SAL':
                try:
                    huadongtaizhaou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '郴州凯程医药有限公司' and fname[5:8] == 'SAL':
                try:
                    kaichen(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '浙江省大德药业集团浙江省医药有限公司' and fname[5:8] == 'SAL':
                try:
                    zhejiangdade(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '金华市东阳医药药材有限公司' and fname[5:8] == 'SAL':
                try:
                    jinhuadongyang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '国药控股海南省鸿益有限公司' and fname[5:8] == 'SAL':
                try:
                    hainanhongyi(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '福建省中源医药有限公司' and fname[5:8] == 'SAL':
                try:
                    fujianzhongyuan(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '四川省本草堂药业有限公司' and fname[5:8] == 'SAL':
                try:
                    sichuanbencao(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '陕西省天士力医药有限公司' and fname[5:8] == 'SAL':
                try:
                    shanxitianlishi(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '江苏宏康医药有限责任公司' and fname[5:8] == 'SAL':
                try:
                    jiangsuhongkang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '山东康诺盛世医药有限公司' and fname[5:8] == 'SAL':
                try:
                    kangnuoshengshi(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '江西省五洲医药营销有限公司' and fname[5:8] == 'SAL':
                try:
                    jiangxiwuzhou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '广西华泰药业有限公司' and fname[5:8] == 'SAL':
                try:
                    guangxihuatai(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '江苏吴中医药销售有限公司' and fname[5:8] == 'SAL':
                try:
                    jiangsuwuzhong(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '重庆市长圣医药有限公司' and fname[5:8] == 'SAL':
                try:
                    chongqingchangsheng(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '湖北人福医药集团有限公司' and fname[5:8] == 'SAL':
                try:
                    hubeirenfu(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '江西省南华医药有限公司' and fname[5:8] == 'SAL':
                try:
                    jiangxihuanan(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '上药科园信海医药湖北有限公司' and fname[5:8] == 'SAL':
                try:
                    keyuanxinhai(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '云南医药工业销售有限公司' and fname[5:8] == 'SAL':
                try:
                    yunnanyiyaogongye(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '武威神洲医药有限责任公司' and fname[5:8] == 'PUR':
                try:
                    wuweishenzhou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '四川腾龙医药有限公司' and fname[5:8] == 'SAL':
                try:
                    sichuantenglong(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '北京市科园信海医药经营有限公司' and fname[5:8] == 'SAL':
                try:
                    beijingkeyuan(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '厦门片仔癀宏仁医药有限公司' and fname[5:8] == 'SAL':
                try:
                    pianzaihaunghongren(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '国药集团西南医药有限公司' and fname[5:8] == 'SAL':
                try:
                    guoyaojituanxinan(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '江西省广力药业有限公司' and fname[5:8] == 'SAL':
                try:
                    jiangxiguangli(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '江西吉安医药有限公司' and fname[5:8] == 'SAL':
                try:
                    jiangxijian(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '山西省康美徕医药有限公司' and fname[5:8] == 'SAL':
                try:
                    shanxikangmeilai(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '浙江省恩泽医药有限公司' and fname[5:8] == 'SAL':
                try:
                    zhejiangenze(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '山东省容大医药有限公司' and fname[5:8] == 'SAL':
                try:
                    shandongronda(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '河北润祥医药有限责任公司' and fname[5:8] == 'SAL':
                try:
                    hebeirunxiang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '菏泽海王医药有限公司' and fname[5:8] == 'SAL':
                try:
                    hezehaiwang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '兰州市强生医药有限责任公司' and fname[5:8] == 'PUR':
                try:
                    lanzhouqiangsheng(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '梅州市卫发医药有限公司' and fname[5:8] == 'PUR':
                try:
                    meizhouweifa(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '南宁华御堂医药公司' and fname[5:8] == 'PUR':
                try:
                    nanninghuayutang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '安徽东升医药物流有限公司' and fname[5:8] == 'INV':
                try:
                    anhuidongshen(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '湖南博瑞药业有限公司' and fname[5:8] == 'INV':
                try:
                    hunanborui(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '广西桂林汇通药业有限公司' and fname[5:8] == 'PUR':
                try:
                    guilinhuitong(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '浙江广为医药有限公司' and fname[5:8] == 'SAL':
                try:
                    zhejiangguangwei(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '浙江广为医药有限公司' and fname[5:8]=='INV':
                try:
                    zhejiangguangwei(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '邢台万邦医药有限责任公司' and fname[5:8] == 'SAL':
                try:
                    xingtaiwanbang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '邢台万邦医药有限责任公司' and fname[5:8] == 'PUR':
                try:
                    xingtaiwanbang_pur(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '安徽省亳州市医药供销有限公司' and fname[5:8] == 'SAL':
                try:
                    anhuihaozhou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '华润芜湖医药有限公司' and fname[5:8] == 'SAL':
                try:
                    huarunwuhu(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '安徽省华源医药股份有限公司' and fname[5:8] == 'SAL':
                try:
                    anhuihuayuan(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '上药控股江苏股份有限公司' and fname[5:8] == 'SAL':
                try:
                    shangyaojiangsu(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '国药控股河南上蔡分公司' and fname[5:8] == 'SAL':
                try:
                    henanshangcai(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '广汉市吉昌药业有限公司' and fname[5:8] == 'SAL':
                try:
                    guanghanjichang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '东辽县医药药材有限责任公司' and fname[5:8] == 'SAL':
                try:
                    dongliaoyiyao(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '淮滨县天一药品经营有限公司' and fname[5:8] == 'SAL':
                try:
                    huaibintianyi(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '陕西省恒庆医药有限公司' and fname[5:8] == 'SAL':
                try:
                    shanxihengqing(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '湖北省诚民天济药业有限公司' and fname[5:8] == 'SAL':
                try:
                    hubeichengmin(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '黑龙江华辰大药房连锁有限公司' and fname[5:8] == 'SAL':
                try:
                    heilongjianghuachen(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '人福医药钟祥有限公司' and fname[5:8] == 'SAL':
                try:
                    renfuzhongxiang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '华润周口医药有限责任公司' and fname[5:8] == 'SAL':
                try:
                    huarunzhoukou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '华润洛阳医药有限责任公司' and fname[5:8] == 'SAL':
                try:
                    huarunluoyang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '华润三门峡医药有限责任公司' and fname[5:8] == 'SAL':
                try:
                    huarunsanmenxia(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '衡水市龙马医药贸易有限公司' and fname[5:8] == 'SAL':
                try:
                    hengshuilongma(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '福建宏海药业有限公司' and fname[5:8] == 'SAL':
                try:
                    fujianhonghai(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '河南九州通国华医药物流有限公司济源分公司' and fname[5:8] == 'SAL':
                try:
                    guohuaqiyuan(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '人福医药天门有限公司' and fname[5:8] == 'SAL':
                try:
                    renfutianmen(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '陕西铭川医药有限公司' and fname[5:8] == 'SAL':
                try:
                    shanximingchuan(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '三台县天诚医药有限公司' and fname[5:8] == 'SAL':
                try:
                    santaitiancheng(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '无锡山禾集团健康参药连锁有限公司' and fname[5:8] == 'SAL':
                try:
                    shanhejiankang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '重庆医药集团周口有限公司' and fname[5:8] == 'SAL':
                try:
                    chongqingzhoukou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '舟山市存德医药有限公司' and fname[5:8] == 'SAL':
                try:
                    zhoushancunde(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '陕西省医药孙思邈五星大药房连锁有限公司' and fname[5:8] == 'SAL':
                try:
                    shanxisunsimiao(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '广西钦州市卫贸发展公司' and fname[5:8] == 'SAL':
                try:
                    qinzhoushiweimao(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '国药控股邓州有限公司' and fname[5:8] == 'SAL':
                try:
                    guokongdengzhou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '吉林省众鑫药业有限公司' and fname[5:8] == 'SAL':
                try:
                    jilinzhongxin(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '江苏致和堂医药物流有限公司' and fname[5:8] == 'SAL':
                try:
                    jiangsuzhihetang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '威海市海王医药有限公司' and fname[5:8] == 'SAL':
                try:
                    weihaihaiwang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '河南省同和堂医药有限公司' and fname[5:8] == 'SAL':
                try:
                    henantongrentang(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            elif result[0] == '广西壮族自治区柳州市医药有限责任公司' and fname[5:8] == 'INV':
                try:
                    guangxiliuzhou(file_path)
                except:
                    print(f'处理失败->|{file_path}')
                    make_floder(data_path, file_path, fname)
            # elif result[0] == '广西壮族自治区柳州市医药有限责任公司' and fname[5:8] == 'PUR':
            #     try:
            #         guangxiliuzhou_pur(file_path)
            #     except:
            #         print(f'处理失败->|{file_path}')
            #         make_floder(data_path, file_path, fname)



            else:
                pass








