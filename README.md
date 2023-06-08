# 👀BigVlookup





## 🐼创意特点

- 以特征向量为 👉🏼**查找值**
- 以数据的每一列为 👉🏼**查找区域**
- 返回相似度最高的列，即👉🏼**列索引**
- So，我称它为 **🤜🏼Big Vlookup🤛🏼** 更为贴切一些
- ℹ️此项目仅做学术交流使用，不得另做它用

## 🐶功能简介

- **文件转换**

  - xls转xlsx

  - 筛选出图片文件，CSS文件，.Ds文件等

    

- **经商识别**

  - 通过映射表精准识别经销商

  - 筛选未识别经销商，新增主数据

    

- **数据清洗**

  - 确定表头

  - 处理空值、异常值、合并单元格、合计行等

    

- **TF-IDF模型的应用**

  - 积累必填字段数据，如产品名称，分词、作为特征向量保存

  - 词嵌入：TD-IDF

  - 提取文件每一列，分词、词嵌入

  - 计算特征向量与每个字段的余弦相似度，相似度最大的则匹配产品名称

    

- **计算相似度，返回列索引**

------

## 🦁项目目录结构

```bash
C:.
│  clean_file.py
│  README.md
│  shxy.py
│  shxy2.py
│  shxy_clean_after.py
│  shxy_clean_data.py
│  shxy_inv.py
│  shxy_pur.py
│  shxy_spe.py
│
├─.idea
│  │  .gitignore
│  │  clean_data3.0.iml
│  │  misc.xml
│  │  modules.xml
│  │  other.xml
│  │  workspace.xml
│  │
│  └─inspectionProfiles
│          profiles_settings.xml
│
├─data
│  │  xy模糊识别库.xlsx
│  │
│  ├─接收文件
│  └─数据备份
├─features
│  └─shxy
│          batch_num.txt
│          factory_manu.txt
│          header_feature.txt
│          product_name.txt
│          product_spec.txt
│
├─img
│      处理前后.png
│
├─util
│  │  report.py
│  │  tfidf.py
│  │
│  └─__pycache__
│          report.cpython-37.pyc
│          report.cpython-39.pyc
│          tfidf.cpython-37.pyc
│          tfidf.cpython-39.pyc
│
└─__pycache__
        shxy_inv.cpython-37.pyc
        shxy_inv.cpython-39.pyc
        shxy_pur.cpython-37.pyc
        shxy_pur.cpython-39.pyc
        shxy_spe.cpython-37.pyc
        shxy_spe.cpython-39.pyc
        szrl_spe.cpython-37.pyc
        szrl_spe.cpython-39.pyc
        test.cpython-37.pyc
```







#### Loading...
