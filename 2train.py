import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.metrics import accuracy_score
from sklearn.multiclass import OneVsRestClassifier
from sklearn.pipeline import Pipeline


if __name__ == '__main__':
    path = 'D:\\garfield\\news-sentiment-detect\\data\\'         # 数据集路径，根据你们自己的修改
    df = pd.read_excel(path + 'train_data_final.xlsx')                      # 读取数据集
    # print(df['sentiment_1'].value_counts())
    # print(df['sentiment_2'].value_counts())
    lens = df.news_content_translate_cn.str.len()                                        # 统计新闻正文长度
    lens.hist(bins=np.arange(0, 5000, 50))
    plt.show()
    # 一共七类
    sentiment_vocab = ['agreeable', 'believable', 'good', 'sad', 'hated', 'worried', 'objective']
    df = pd.concat([df, pd.DataFrame(columns=sentiment_vocab)], axis=1)
    df = df.fillna(0)           # 若有缺失值，自动填充为0
    # 构造数据集格式，类别标签一共七列，0代表不属于该类，1代表属于
    # 详情可见print(df)输出
    for i in range(len(df)):
        sentiment_1 = df['news_emotion.0'].loc[i]
        sentiment_2 = df['news_emotion.1'].loc[i]
        if sentiment_1 in sentiment_vocab:
            df[sentiment_1].loc[i] = 1
        if sentiment_2 in sentiment_vocab:
            df[sentiment_2].loc[i] = 1
    print(df)
    # 划分数据集-> 85%作为训练集，15%作为测试集
    # 训练集用于训练情感分类模型，测试集用于测试训练后的模型精度
    train, test = train_test_split(df, random_state=42, test_size=0.15, shuffle=True)

    X_train = train.news_content_translate_cn            # 训练数据为新闻正文
    X_test = test.news_content_translate_cn              # 测试数据为新闻正文

    # Define a pipeline combining a text feature extractor with multi lable classifier
    NB_pipeline = Pipeline([
        ('tfidf', TfidfVectorizer(ngram_range=(1, 2), min_df=3, max_df=0.9, strip_accents='unicode', use_idf=1)),
        ('clf', OneVsRestClassifier(MultinomialNB(
            fit_prior=True, class_prior=None))),
    ])
    # for循环，依次训练这七个情感分类器
    for category in sentiment_vocab:
        print('... Processing {}'.format(category))
        NB_pipeline.fit(X_train, train[category])       # 训练
        prediction = NB_pipeline.predict(X_test)        # 预测
        # print(prediction)
        # compute the testing accuracy
        print('Test accuracy is {}'.format(accuracy_score(test[category], prediction)))
        # 预测两个新闻标题
        result = NB_pipeline.predict_proba(['Detection of asymptomatic cases coronavirus and pass alert experts',
                                            'Map of the expansion of the coronavirus'])
        print(result)       # 第一个数代表是概率，第二个代表不是的概率