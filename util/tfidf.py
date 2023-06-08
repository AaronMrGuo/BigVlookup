import jieba
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer


class CleanModel(object):
    """
    :param file_path:要提取的文件路径;
    :param features:提取数据的特征;
    """

    def __init__(self, file_path, features):
        df = pd.read_excel(file_path, dtype="object")
        self.df = df
        self.features = features

    def model(self):
        docs = []
        for i in range(self.df.shape[1]):
            sens = str()
            for j in self.df.iloc[:, i]:
                sens = str(j) + ' ' + sens
            tokens = list(set(jieba.cut(sens, cut_all=False)))
            token = str()
            for x in tokens:
                token = x + ' ' + token
            docs.append(token)
        docs.append(self.features)
        vectorizer = TfidfVectorizer()
        model = vectorizer.fit_transform(docs)
        tfidf = model.todense().round(6)
        cos_sims = []
        for i in range(len(tfidf) - 1):
            values = tfidf[-1]
            cos_sim = (np.dot(tfidf[i], values) / (np.linalg.norm(tfidf) * np.linalg.norm(values) + 1)).round(6)
            cos_sims.append(cos_sim)
        cos_max_sim = np.max(np.array(cos_sims)).round(6)
        columns_index = cos_sims.index(cos_max_sim)
        return columns_index, cos_max_sim
