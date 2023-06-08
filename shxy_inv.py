import pandas as pd
import numpy as np
import os, re, time, datetime, shutil
import jieba
from sklearn.feature_extraction.text import TfidfVectorizer
import warnings

warnings.filterwarnings('ignore')


# 黑名单经销商


# '河南省康信医药有限公司':'库存流向： 有效期转换错误',
# '国药控股湖北宏源医药有限公司':'库存流向： 有效期 为数值,转换错误',
# '国药控股湖北省有限公司':'库存流向： 有效期 为数值,转换错误',
# '江苏吴中医药销售有限公司':'库存流向：效期 格式错误',
def get_spe_inv(data_path):
    print('xy 正在获取黑名单经销商，请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    spe_name = {
        '邢台天宇医药有限公司': '库存流向：原因待核实',
        '商丘市华杰医药有限公司': '库存流向：原因待核实',
        '江苏宏康医药有限责任公司': '库存流向：表头错误',
        '广西华泰药业有限公司': '库存流向：产品名称不全',
        '浙江长典医药有限公司': '库存流向：产品名称 规格 生产厂家 在一列',
        '舟山市普陀区医药药材有限公司': '库存流向：产品名称 规格 生产厂家 在一列',
        '国药控股吉林省有限公司': '库存流向：有效期错列',
        '福建省雷允上医药有限公司': '库存流向：产品名称 规格 生产厂家 在一列',
        '温州市时代医药有限公司': '库存流向： 数据跑完缺失了',
        '安徽东升医药物流有限公司': '库存流向： 原始库存文件，日期降序，结存数量升序，产品相关的几个字段联立，保留排序第一的数据，其他删除',
        '国药控股河南省股份有限公司': '库存流向： 含分子公司，只保留总公司',
        '黑龙江省国龙医药有限公司': '库存流向： 人工处理'
        # '广西壮族自治区柳州市医药有限责任公司':'库存流向：需要按照产品名称，规格，批号区分',
    }
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'INV':
                dps_name = file_path.split('\\')[-2].split('_')[1]
                for i in spe_name:
                    if dps_name == i:
                        floder_name = data_path + '处理失败' + '\\' + '黑名单经销商' + '\\' + '\\'.join(
                            file_path.split('\\')[len(data_path.split('\\')):-1])
                        if not os.path.exists(floder_name):
                            os.makedirs(floder_name)
                        new_file_path = floder_name + '\\' + fname
                        shutil.move(file_path, new_file_path)
                        print(f'xy：黑名单经销商->|{file_path}')
                        print(f'原因: {spe_name[i]}')
                        print('-' * 200)
                # else:
                #     print(f'xy：黑名单经销商，文件命名错误->|{file_path}')
                #     print('-'*200)


# 库存日期
def parse_date_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            # print(file_path)
            if fname[5:8] == 'INV':
                pattern = re.compile('\d{8}')
                inv_date = pattern.findall(fname)
                if len(inv_date) == 1:
                    df = pd.read_excel(file_path, dtype='object')
                    # if '库存日期' not in list(df.columns):
                    df['库存日期'] = inv_date[0]
                    try:
                        df['库存日期'] = df['库存日期'].map(lambda x: str(pd.to_datetime(x).date())[:10])
                        df.to_excel(file_path, index=False)
                    except:
                        floder_name = data_path + '处理失败' + '\\' + '库存日期' + '\\' + '\\'.join(
                            file_path.split('\\')[len(data_path.split('\\')):-1])
                        if not os.path.exists(floder_name):
                            os.makedirs(floder_name)
                        new_file_path = floder_name + '\\' + fname
                        shutil.move(file_path, new_file_path)
                else:
                    floder_name = data_path + '处理失败' + '\\' + '库存日期' + '\\' + '\\'.join(
                        file_path.split('\\')[len(data_path.split('\\')):-1])
                    if not os.path.exists(floder_name):
                        os.makedirs(floder_name)
                    new_file_path = floder_name + '\\' + fname
                    shutil.move(file_path, new_file_path)


def inv_num_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if '数量' not in list(df.columns):
                    for i in df.columns:
                        if i == '总库存':
                            df.rename(columns={'总库存': '数量'}, inplace=True)
                            df.to_excel(file_path, index=False)
                            break
                        elif i == '可销数量':
                            df.rename(columns={'可销数量': '数量'}, inplace=True)
                            df.to_excel(file_path, index=False)
                            break
                        elif i == '库存总数':
                            df.rename(columns={'库存总数': '数量'}, inplace=True)
                            df.to_excel(file_path, index=False)
                            break
                        elif i == '总数量':
                            df.rename(columns={'总数量': '数量'}, inplace=True)
                            df.to_excel(file_path, index=False)
                            break
                        elif i == '库存数量':
                            df.rename(columns={'库存数量': '数量'}, inplace=True)
                            df.to_excel(file_path, index=False)
                            break
                        elif i == '可用量':
                            df.rename(columns={'可用量': '数量'}, inplace=True)
                            df.to_excel(file_path, index=False)
                            break
                        elif i == '总库存数':
                            df.rename(columns={'总库存数': '数量'}, inplace=True)
                            df.to_excel(file_path, index=False)
                            break


def get_product_num(data_path):
    print('xy库存：正在提取 库存数量 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    maping_num = ['数量', '药品数量', '开票数量', '库管数量', '合格数量', '库存总数', '实际库存', '库存数量', '可销库存', '可售库存', '数量（盒', '库存数',
                  '可分配数量', '库存量', '库存数量', '库存小计', '总数量', '库 存', '库 存', '库存总数量', '总库存', '结存数量', '可用数量', '货品数量',
                  '当前库存', '库存', '基本单位数量', '批号数量', '当前库存量', '参考结余', '非限制', '帐存数', '可销库存', '财务库存数量', '可销数量', '商品数量',
                  '当前库存数', '仓库库存数量', '实际数量', '合格仓', '在库库存数量', '实际库存数量', '数量[汇总][汇总', '仓库库存数量汇总', '仓库数量', '结存数',
                  '可用库存', '数量[汇总', '库存数量[汇总', '可销数', '库房数量', '数量（瓶', '库余']
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0 and '数量' not in list(df.columns):
                    list1 = df.columns
                    column_names = [column_name for column_name in list1 if column_name in maping_num]
                    if len(column_names) == 0:
                        floder_name = data_path + '处理失败' + '\\' + '库存数量' + '\\' + '\\'.join(
                            file_path.split('\\')[len(data_path.split('\\')):-1])
                        if not os.path.exists(floder_name):
                            os.makedirs(floder_name)
                        new_file_path = floder_name + '\\' + fname
                        shutil.move(file_path, new_file_path)
                    elif len(column_names) == 1 and column_names[0] != '批号':
                        df['数量'] = df.loc[:, column_names[0]]
                        #                 df.rename(columns={column_names[0]:'数量'},inplace=True)
                        df.to_excel(file_path, index=False)
                    elif len(column_names) > 1:
                        if '数量' not in column_names:
                            print(column_names)
                            print(file_path)
                            print('=' * 200)
                            msg = str(input('请输入数量字段 输入N手动处理 '))
                            if msg == 'N':
                                floder_name = data_path + '处理失败' + '\\' + '库存数量' + '\\' + '\\'.join(
                                    file_path.split('\\')[len(data_path.split('\\')):-1])
                                if not os.path.exists(floder_name):
                                    os.makedirs(floder_name)
                                new_file_path = floder_name + '\\' + fname
                                shutil.move(file_path, new_file_path)
                            else:
                                df['数量'] = df.loc[:, msg]
                                #                     df.rename(columns={msg:'数量'},inplace=True)
                                df.to_excel(file_path, index=False)
                    else:
                        print(f'空文件-> {file_path}')


# 产品名称
def product_name_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                for i in df.columns:
                    if i == '产品名称':
                        break
                    elif i == '商品名称' or i == '药品名称' or i == '货品名称' or i == '药品信息' or i == '药品通用名称' \
                            or i == '品名规格' or i == '品名\规格' or i == '通用名(商品名)\规 格':
                        df.rename(columns={i: '产品名称'}, inplace=True)
                        df.to_excel(file_path, index=False)
                        break


def tfidf(df, features):
    global columns_index, cos_sim_max
    docs = []
    for i in range(df.shape[1]):
        sens = str()
        for j in df.iloc[:, i]:
            sens = str(j) + ' ' + sens
        tokens = list(set(jieba.cut(sens, cut_all=False)))
        token = str()
        for x in tokens:
            token = x + ' ' + token
        docs.append(token)
    docs.append(features)
    vectorizer = TfidfVectorizer()
    model = vectorizer.fit_transform(docs)
    tfidf = model.todense().round(6)
    return tfidf


def column_index(tfidf):
    cos_sims = []
    for i in range(len(tfidf) - 1):
        values = tfidf[-1]
        cos_sim = (np.dot(tfidf[i], values) / (np.linalg.norm(tfidf) * np.linalg.norm(values) + 1)).round(6)
        cos_sims.append(cos_sim)
    cos_max_sim = np.max(np.array(cos_sims)).round(6)
    # print(cos_sims)
    columns_index = cos_sims.index(cos_max_sim)
    return columns_index, cos_max_sim


# xy：产品名称
def get_product_name(data_path):
    print('xy库存：正在提取 产品名称 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    product_name_list = []
    with open(r'.\features\shxy\product_name.txt') as f:
        product_name_feature = f.readlines()[0]
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0:
                    if '产品名称' not in list(df.columns):
                        # docs = get_docs(df, product_name_feature)
                        tfidf_value = tfidf(df, product_name_feature)
                        product_name_index, cos_sim_max = column_index(tfidf_value)
                        product_name_list.append(cos_sim_max)
                        df = pd.read_excel(file_path, dtype='object')
                        if cos_sim_max >= 0.01:  # cos_sim均值0.1645
                            df['产品名称'] = df.iloc[:, product_name_index]
                            df.to_excel(file_path, index=False)
                        else:
                            floder_name = data_path + '处理失败' + '\\' + '产品名称' + '\\' + '\\'.join(
                                file_path.split('\\')[len(data_path.split('\\')):-1])
                            if not os.path.exists(floder_name):
                                os.makedirs(floder_name)
                            new_file_path = floder_name + '\\' + fname
                            shutil.move(file_path, new_file_path)
                            print(f'xy库存 产品名称 提取失败->|{new_file_path}')
                            print('-' * 200)


# xy：规格
def product_spe_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                for i in df.columns:
                    if i == '规格':
                        break
                    elif i == '商品规格' or i == '药品规格' or i == '品种规格' or i == '产品规格' or i == '货品规格' or i == '规格/型号' or i == '规格型号' \
                            or i == '包装规格' or i == '商品规格/型号':
                        df.rename(columns={i: '规格'}, inplace=True)
                        df.to_excel(file_path, index=False)


def get_product_spe(data_path):
    print('xy库存：正在提取 规格 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    product_spe_list = []
    with open(r'.\features\shxy\product_spec.txt') as f:
        product_spe = f.readlines()[0]
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0:
                    if '规格' not in list(df.columns):
                        # docs = get_docs(df, product_spe)
                        tfidf_value = tfidf(df, product_spe)
                        product_spe_index, cos_sim_max = column_index(tfidf_value)
                        product_spe_list.append(cos_sim_max)
                        df = pd.read_excel(file_path, dtype='object')
                        if cos_sim_max >= 0.04:  # cos_sim均值0.1645
                            df['规格'] = df.iloc[:, product_spe_index]
                            df.to_excel(file_path, index=False)
                        else:
                            floder_name = data_path + '处理失败' + '\\' + '规格' + '\\' + '\\'.join(
                                file_path.split('\\')[len(data_path.split('\\')):-1])
                            if not os.path.exists(floder_name):
                                os.makedirs(floder_name)
                            new_file_path = floder_name + '\\' + fname
                            shutil.move(file_path, new_file_path)
                            print(f'xy库存 规格 提取失败->|{new_file_path}')
                            print('-' * 200)


# xy：生产厂家
def get_product_manu(data_path):
    print('xy库存：正在提取 生产厂家 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    with open(r'.\features\shxy\factory_manu.txt') as f:
        product_manu = f.readlines()[0]
    factory_name_list = []
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0:
                    if '生产厂家' not in list(df.columns):
                        # docs = get_docs(df, product_manu)
                        tfidf_value = tfidf(df, product_manu)
                        product_manu_index, cos_sim_max = column_index(tfidf_value)
                        factory_name_list.append(cos_sim_max)
                        df = pd.read_excel(file_path, dtype='object')
                        if cos_sim_max >= 0.029:  # cos_sim均值0.1645
                            df['生产厂家'] = df.iloc[:, product_manu_index]
                            df.to_excel(file_path, index=False)
                        else:
                            # print(cos_sim_max)
                            # print(file_path)
                            df['生产厂家'] = 'xy'
                            df.to_excel(file_path, index=False)


# xy：批号

def batch_num_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                for i in df.columns:
                    if i == '批号':
                        break
                    elif i == '批号信息' or i == '批号效期' or i == '商品批号/效期' or i == '生产批号' or i == '订单批号' or i == '产品批号' or i == '商品批号' \
                            or i == '批号/序列号' or i == '销售批号' or i == '药品批号' or i == '货品批号' or i == '商品来源批号':
                        df.rename(columns={i: '批号'}, inplace=True)
                        df.to_excel(file_path, index=False)


def get_batch_num(data_path):
    print('xy库存：正在提取 批号 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    batch_num_list = []
    with open(r'.\features\shxy\batch_num.txt') as f:
        batch_num = f.readlines()[0]
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0:
                    if '批号' not in list(df.columns):
                        # docs = get_docs(df, batch_num)
                        tfidf_value = tfidf(df, batch_num)
                        batch_num_index, cos_sim_max = column_index(tfidf_value)
                        batch_num_list.append(cos_sim_max)
                        df = pd.read_excel(file_path, dtype='object')
                        if cos_sim_max >= 0.04:  # cos_sim均值0.1645
                            df['批号'] = df.iloc[:, batch_num_index]
                            df.to_excel(file_path, index=False)
                        else:

                            # floder_name = data_path + '处理失败' + '\\' + '批号' + '\\' + '\\'.join(
                            # file_path.split('\\')[len(data_path.split('\\')):-1])
                            # if not os.path.exists(floder_name):
                            # os.makedirs(floder_name)
                            # new_file_path = floder_name + '\\' + fname
                            # shutil.move(file_path, new_file_path)

                            print(f'xy库存 批号 提取失败->|{file_path}')
                            print('-' * 200)


# xy：产品单位
def get_product_unit(data_path):
    print('xy库存：正在提取 产品单位 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                product_unit_count = 0
                list1 = []
                if df.shape[0] > 0:
                    for i in range(df.shape[1]):
                        for j in df.iloc[:, i]:
                            j = str(j).replace(' ', '')
                            if j == '支' or j == '瓶' or j == '盒' or j == '包' or j == '套' or j == 'KG' or j == '袋' or j == '听':
                                product_unit_count += 1
                        list1.append(product_unit_count)
                    if product_unit_count / df.shape[0] >= 0.5:
                        product_unit_index = list1.index(np.max(np.array(list1)))
                        df.rename(columns={list(df.columns)[product_unit_index]: '产品单位'}, inplace=True)
                        df.to_excel(file_path, index=False)
                    #             df['产品单位']=df.iloc[:,product_unit_index]
                    else:
                        df['产品单位'] = '缺失'
                        df.to_excel(file_path, index=False)


# xy：单价
def get_product_price(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                for i in df.columns:
                    if '单价' in i or '零售价' in i or '售价' in i:
                        price_index = list(df.columns).index(i)
                        df['单价'] = df.iloc[:, price_index]
                        df.to_excel(file_path, index=False)


# xy：金额
def get_product_amount(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                for i in df.columns:
                    if '金额' in i:
                        amount_index = list(df.columns).index(i)
                        df['金额'] = df.iloc[:, amount_index]
                        df.to_excel(file_path, index=False)


# xy： 库存 有效期  修改
def deter_date_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                for i in df.columns:
                    if i == '有效期':
                        df.rename(columns={'有效期': '效期'}, inplace=True)
                        # T_date(df)
                        df.to_excel(file_path, index=False)
                    elif i == '失效日期':
                        df.rename(columns={'失效日期': '效期'}, inplace=True)
                        # T_date(df)
                        df.to_excel(file_path, index=False)


# xy： 库存 有效期
def get_deter_date(data_path):
    print('xy库存：正在提取 有效期 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    mapping_custer_name = ['有效期至', '有效期', '失效日期', '有效日期', '保质日期', '保质期至', '失效期', '商品有效期', \
                           '有效期至/失效日期', '药品有效期至/医疗器械失效日期', '时效日期', '有效期限', '有效月',
                           '有效期(月)', '有效期(月', '灭菌有效期']
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0 and '效期' in list(df.columns):
                    break
                else:
                    list1 = df.columns
                    column_names = [column_name for column_name in list1 if column_name in mapping_custer_name]
                    if len(column_names) == 0:

                        pass
                        # print(df.columns)
                        # print(file_path)
                        # msg=str(input('请添加 有效期 映射字段,输入C继续,输入N手动处理 '))
                        # if msg == 'N':
                        #     floder_name = data_path + '处理失败' + '\\' + '有效期' + '\\' + '\\'.join(
                        #         file_path.split('\\')[len(data_path.split('\\')):-1])
                        #     if not os.path.exists(floder_name):
                        #         os.makedirs(floder_name)
                        #     new_file_path = floder_name + '\\' + fname
                        #     shutil.move(file_path, new_file_path)
                        # elif msg == 'C':
                        #     continue
                    #         else:
                    #             df['有效期']=df.loc[:,msg]
                    # #                 df.rename(columns={msg:'客户名称'},inplace=True)
                    #             df.to_excel(file_path,index=False)
                    elif len(column_names) == 1 and column_names[0] != '效期':
                        df['效期'] = df.loc[:, column_names[0]]
                        #                 df.rename(columns={column_names[0]:'客户名称'},inplace=True)
                        df.to_excel(file_path, index=False)
                    elif len(column_names) > 1:
                        if '效期' not in column_names:
                            print(column_names)
                            print(file_path)
                            print('=' * 200)
                            msg = str(input('请输入 效期 字段,输入N手动处理 '))
                            if msg == 'N':
                                floder_name = data_path + '处理失败' + '\\' + '效期' + '\\' + '\\'.join(
                                    file_path.split('\\')[len(data_path.split('\\')):-1])
                                if not os.path.exists(floder_name):
                                    os.makedirs(floder_name)
                                new_file_path = floder_name + '\\' + fname
                                shutil.move(file_path, new_file_path)
                            else:
                                df.rename(columns={msg: '效期'}, inplace=True)
                                df.to_excel(file_path, index=False)

                # else:
                #     print(f'空文件-> {file_path}')


# xy：产品编码
def get_product_id(data_path):
    print('xy库存：正在提取 产品编码 ,请稍后...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    maping_num = ['商品编码', '商品编号', '商品编号', '货品明细ID', '品种编码', '规格/品规ID', '新商品编码', '药品编码', '产品编号', '商品ID', '品种号',
                  '货品编号', '货号', '物料编码', '商品主编码', '货品编码 / 商品编码', '货品编码', '药品编号', '药品M码', '商品代码']
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0:
                    list1 = df.columns
                    column_names = [column_name for column_name in list1 if column_name in maping_num]
                    if len(column_names) == 0:
                        pass
                    #             print(df.columns)
                    #             print(file_path)
                    #             msg=str(input('请添加 产品编码 映射字段 '))
                    #             df['产品编码']=df.loc[:,msg]
                    # #                 df.rename(columns={msg:'数量'},inplace=True)
                    #             df.to_excel(file_path,index=False)
                    elif len(column_names) == 1 and column_names[0] != '产品编码':
                        df['产品编码'] = df.loc[:, column_names[0]]
                        #                 df.rename(columns={column_names[0]:'数量'},inplace=True)
                        df.to_excel(file_path, index=False)
                    elif len(column_names) > 1:
                        pass
                    #             if '产品编码' not in column_names:
                    #                 print(column_names)
                    #                 print(file_path)
                    #                 print('='*130)
                    #                 msg=str(input('请输入 产品编码 字段 '))
                    #                 df['产品编码']=df.loc[:,msg]
                    # #                     df.rename(columns={msg:'数量'},inplace=True)
                    #                 df.to_excel(file_path,index=False)
                    else:
                        print(f'空文件-> {file_path}')


# 时间格式转化
def T_date(df):
    for i in df.columns:
        if df[i].dtype == 'datetime64[ns]':
            df[i] = df[i].apply(lambda x: str(pd.to_datetime(x).date()))


def check_key(data_path):
    print('正在检查 关键字段是否缺失...')
    print('=' * 200)
    new_data_path = data_path + '处理成功'
    list1 = ['库存日期', '产品名称', '规格', '产品单位', '数量', '生产厂家']
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'INV':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path, dtype='object')
                if df.shape[0] > 0:
                    for i in list1:
                        if i not in list(df.columns):
                            new_floder = data_path + '处理失败' + '\\' + str(i) + '\\' + '\\'.join(
                                file_path.split('\\')[-2:-1])
                            if not os.path.exists(new_floder):
                                os.makedirs(new_floder)
                            new_file_path = os.path.join(new_floder, fname)
                            shutil.move(file_path, new_file_path)
                            print(f'{i} 字段缺失->| {file_path}')
                            print('-' * 200)
                else:
                    floder_name = data_path + '处理失败' + '\\' + '数据缺失' + '\\' + '\\'.join(
                        file_path.split('\\')[len(data_path.split('\\')):-1])
                    if not os.path.exists(floder_name):
                        os.makedirs(floder_name)
                    new_file_path = floder_name + '\\' + fname
                    shutil.move(file_path, new_file_path)

    print('xy：关键字段检查完成，准备清洗数据...')
    print('=' * 200)
    # print('=' * 200)


# 去除非必填字段，选填字段
def reduce_data(data_path):
    new_data_path = data_path + '处理成功'
    df_final = pd.DataFrame()
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'INV':
                # print(file_path)
                df = pd.read_excel(file_path, dtype='object')
                list1 = ['库存日期', '产品名称', '规格', '产品单位', '数量', '批号', '单价', '金额', '生产厂家', '效期']
                list2 = list(df.columns)
                column_names = [column_name for column_name in list1 if column_name in list2]
                df_final = df[column_names]
                # df.dropna(how='any',axis=1,inplace=True)
                df_final = df_final[
                    (df_final['数量'].notnull()) & (df_final['数量'] != '合计：') & (df_final['数量'] != '&nbsp;')]
                T_date(df_final)
                df_final.to_excel(file_path, index=False)


def del_flows(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'INV':
                # print(file_path)
                df = pd.read_excel(file_path, dtype='object')
                if '效期' in list(df.columns):
                    df = df[(df['效期'] != '合计：')]
                    # df = df[(df['库存日期'].notnull()) & (df['库存日期'] != '合计') & (df['库存日期'] != '合计：') & (
                    #             df['库存日期'] != 'NaT') & (df['库存日期'] != '业务日期')]
                    df = df[(df['产品单位'].notnull()) & (df['产品单位'] != '----------')]
                    df = df[(df['产品名称'] != '合计') & (df['产品名称'] != '小计') & (df['产品名称'] != '----------') & (
                        df['产品名称'].notnull())]

                    # df = df[(df['客户名称'].notnull()) & (df['客户名称'] != '~')]
                    df.to_excel(file_path, index=False)
                    # else:
                    df = df[(df['库存日期'].notnull()) & (df['库存日期'] != '合计') & (df['库存日期'] != '合计：') & (
                                df['库存日期'] != 'NaT') & (df['库存日期'] != '业务日期')]
                    df = df[(df['产品单位'].notnull()) & (df['产品单位'] != '----------')]
                    df = df[(df['产品名称'] != '合计') & (df['产品名称'] != '----------')]

                    # df = df[(df['客户名称'].notnull()) & (df['客户名称'] != '~')]
                    df.to_excel(file_path, index=False)
                else:
                    df = df[(df['产品单位'].notnull()) & (df['产品单位'] != '----------')]
                    df = df[(df['产品名称'] != '合计') & (df['产品名称'] != '小计') & (df['产品名称'] != '----------') & (
                        df['产品名称'].notnull())]

                    # df = df[(df['客户名称'].notnull()) & (df['客户名称'] != '~')]
                    df.to_excel(file_path, index=False)
                    # else:
                    df = df[(df['库存日期'].notnull()) & (df['库存日期'] != '合计') & (df['库存日期'] != '合计：') & (
                            df['库存日期'] != 'NaT') & (df['库存日期'] != '业务日期')]
                    df = df[(df['产品单位'].notnull()) & (df['产品单位'] != '----------')]

                    # df = df[(df['客户名称'].notnull()) & (df['客户名称'] != '~')]
                    df.to_excel(file_path, index=False)


def add_factory_manue(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'INV':
                df = pd.read_excel(file_path, dtype='object')
                if '生产厂家' not in list(df.columns):
                    df['生产厂家'] = df['生产厂家'].fillna('xy')
                    df.to_excel(file_path, index=False)


# 有效期转换
def inv_date_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            # print(file_path)
            # df = pd.read_excel(file_path, dtype='object')
            if fname[5:8] == 'INV':
                df = pd.read_excel(file_path, dtype='object')
                if '效期' in list(df.columns):
                    try:
                        df = pd.read_excel(file_path, dtype='object')
                        df['效期'] = df['效期'].map(lambda x: str(pd.to_datetime(x))[:10])
                        # T_date(df)
                        df['效期'] = df['效期'].str.replace('NaT', '')
                        df.to_excel(file_path, index=False)
                    except:
                        print(f'xy库存:效期 清洗失败->| {file_path}')
                        print('-' * 200)

                # df['库存日期']=df['库存日期'].map(lambda x: str(pd.to_datetime(x))[:10])
                # T_date(df)
                # df.to_excel(file_path,index=False)
                # if '有效期' in list(df.columns):
                #     try:
                #         df = pd.read_excel(file_path, dtype='object')
                #         df['有效期'] = df['有效期'].map(lambda x: str(pd.to_datetime(x))[:10])
                #         T_date(df)
                #         df.to_excel(file_path, index=False)
                #     except:
                #         print(f'xy库存:有效期 清洗失败->| {file_path}')
            # elif '有效期' in list(df.columns):
            #     df = pd.read_excel(file_path, dtype='object')
            #     df['有效期'] = df['有效期'].map(lambda x: str(pd.to_datetime(x))[:10])
            #     T_date(df)
            #     df.to_excel(file_path, index=False)


# 批号数据清洗
def inv_batch_num_clean(data_path):
    new_data_path = data_path + '处理成功'
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            if fname[5:8] == 'PUR':
                file_path = os.path.join(dirpath, fname)
                df = pd.read_excel(file_path)
                if df.shape[0] > 1 and '批号' in list(df.columns):
                    df['批号'] = df['批号'].map(lambda x: str(x))
                    df['批号'] = df['批号'].str.extract('([a-zA-Z]\d+|\d+)')
                    df = df.replace('nan', '')
                    df.to_excel(file_path, index=False)


def shxy_inv_excel_style(df, data_path, fname):
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    import itertools, os
    writer = pd.ExcelWriter(os.path.join(data_path, fname), engine='openpyxl')
    df.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    font = Font(name='微软雅黑', bold=True, color='f7f7f7')
    alignment = Alignment(vertical='top', wrap_text=True)
    pattern_fill = PatternFill(fill_type='solid', fgColor='00b0f0')
    side = Side(style='thin')
    border = Border(left=side, right=side, top=side, bottom=side)
    for cell in itertools.chain(*worksheet['A1:K1']):
        cell.font = font
        cell.alignment = alignment
        cell.fill = pattern_fill
        cell.border = border
    worksheet.column_dimensions['A'].width = 12
    worksheet.column_dimensions['B'].width = 20
    worksheet.column_dimensions['C'].width = 15
    worksheet.column_dimensions['D'].width = 10
    worksheet.column_dimensions['E'].width = 8
    worksheet.column_dimensions['F'].width = 10
    worksheet.column_dimensions['G'].width = 8
    worksheet.column_dimensions['H'].width = 8
    worksheet.column_dimensions['I'].width = 35
    worksheet.column_dimensions['J'].width = 12
    worksheet.column_dimensions['K'].width = 35
    # worksheet.column_dimensions['L'].width = 35
    # worksheet.column_dimensions['M'].width = 30
    # worksheet.column_dimensions['N'].width = 12
    # worksheet.column_dimensions['O'].width = 35

    writer.save()
    writer.close()


def check_data(data_path):
    new_data_path = data_path + '处理成功'
    concat_df = pd.DataFrame()
    for dirpath, dirname, filenames in os.walk(new_data_path):
        for fname in filenames:
            file_path = os.path.join(dirpath, fname)
            if fname[5:8] == 'INV':
                df = pd.read_excel(file_path, dtype='object')
                #         print(file_path)
                #         pattern=re.compile('[\u4e00-\u9fa5].*[司站队部房店肃心行院]')
                #         result=pattern.findall(fname)[0]
                df['经销商'] = file_path.split('\\')[-2].split('_')[1]
                concat_df = pd.concat([concat_df, df], axis=0)
    list1 = ['库存日期', '产品名称', '规格', '产品单位', '数量', '批号', '单价', '金额', '生产厂家', '效期', '经销商']
    list2 = list(concat_df.columns)
    column_name = [x for x in list1 if x not in list2]
    for i in column_name:
        concat_df[i] = ''
    concat_df = concat_df[
        ['库存日期', '产品名称', '规格', '产品单位', '数量', '批号', '单价', '金额', '生产厂家', '效期', '经销商']]
    shxy_inv_excel_style(concat_df, '\\'.join(data_path.split('\\')[:-1]), data_path.split('\\')[-1] + '数据合并INV.xlsx')
    # concat_df.to_excel(data_path+'数据合并INV.xlsx', index=False)


# if __name__ == '__main__':
#     data_path=r'C:\Users\guodingyu\Desktop\工具\SHXY_CLEAN\接收文件一级商\20220228'
def inv_clean(data_path):
    time_start = time.time()  # 记录开始时间
    get_spe_inv(data_path)  # 黑名单经销商
    parse_date_clean(data_path)  # 库存日期
    inv_num_clean(data_path)  # 数量
    get_product_num(data_path)  # 数量
    product_name_clean(data_path)  # 修改 产品名称
    get_product_name(data_path)  # 提取 产品名称
    product_spe_clean(data_path)  # 修改  规格
    get_product_spe(data_path)  # 提取 规格
    batch_num_clean(data_path)  # 修改 批号
    get_batch_num(data_path)  # 提取 批号
    get_product_manu(data_path)  # 提取 生产厂家
    get_product_unit(data_path)  # 提取 产品单位
    get_product_price(data_path)  # 提取 单价
    deter_date_clean(data_path)  # 效期修改
    get_deter_date(data_path)  # 提取 有效期
    get_product_id(data_path)  # 提取 产品编号

    check_key(data_path)  # 检查必填字段是否缺失
    # if msg == 1:
    # print('关键字段检查完成，开始清洗数据...')
    reduce_data(data_path)
    inv_date_clean(data_path)  # 有效期转换
    del_flows(data_path)
    add_factory_manue(data_path)
    inv_batch_num_clean(data_path)
    check_data(data_path)
    # else:
    # print('关键字段缺失，请检查...')
    time_end = time.time()
    time_sum = time_end - time_start
    print(f'inv_clean running->| {round(time_sum, 2)}s ')
    print('=' * 200)
