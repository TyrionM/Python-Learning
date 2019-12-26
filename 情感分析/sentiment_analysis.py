import pandas as pd
import jieba
import codecs

"""
1. 文本切割
"""


def wordslist(filepath):
    words_list = [line.strip() for line in codecs.open(filepath, 'r',"utf-8").readlines()]
    return words_list


def stopwordlist(filepath):
    stop_word_list = [line.strip() for line in codecs.open(filepath, 'r', "utf-8").readlines()]
    de_no_list = wordslist("noDict.txt")
    de_degree_list = wordslist("clean_degreeDict.txt")
    for p in range(4):
        for w in stop_word_list:
            if w in de_no_list or w in de_degree_list:
                stop_word_list.remove(w)
    return stop_word_list


def sent2word(sentence):
    """
    Segment a sentence to words
    Delete stopwords
    """
    segList = jieba.cut(sentence)
    segResult = []
    for w in segList:
        segResult.append(w)
    stopwords = stopwordlist('stop_word.txt')
    sent_word_list = []
    for word in segResult:
        if word in stopwords:
            # print "stopword: %s" % word
            continue
        else:
            sent_word_list.append(word)
    return sent_word_list


"""
2. 情感定位
"""
def sentiment_loc_with_score(my_word_list):
    # (1) 情感词字典
    senList = wordslist('BosonNLP_sentiment_score.txt')
    senDict = {}
    for s in senList:
        try:
            senDict[s.split(' ')[0]] = s.split(' ')[1]
        except:
            print(s)
    # (2) 否定词列表
    notList = wordslist('noDict.txt')
    # (3) 程度副词字典
    degreeList = wordslist('degreeDict.txt')
    degreeDict = {}
    for d in degreeList:
        degreeDict[d.split(',')[0]] = d.split(',')[1]
    my_sentiment_dict = {}
    my_no_word_dict = {}
    my_degree_dict = {}
    for i in range(len(my_word_list)):
        word = my_word_list[i]
        if word in senDict and word not in notList and word not in degreeDict:
            my_sentiment_dict[i] = [word,senDict[word]]
        elif word in notList and word not in degreeDict:
            my_no_word_dict[i] = [word,-1]
        elif word in degreeDict:
            my_degree_dict[i] = [word,degreeDict[word]]
    my_group ={}
    g_number = 0
    my_group[0] = []
    for i in range(len(my_word_list)):
        if i in my_sentiment_dict:
            my_group[g_number].append(my_sentiment_dict[i])
            g_number = g_number + 1
            my_group[g_number] = []
        if i in my_no_word_dict:
            my_group[g_number].append(my_no_word_dict[i])
        if i in my_degree_dict:
            my_group[g_number].append(my_degree_dict[i])
    return my_group

def count_score(my_dict):
    my_score = 0

    tag =0
    for d in my_dict:
        if len(my_dict[d]) == 0 :
            tag =tag + 1
        else:
            m = 1
            for s in my_dict[d]:
                m = m*float(s[1])
            my_score = my_score + m
    my_length = len(my_dict) - tag

    return my_score,my_length



if __name__ == "__main__":
    my_dict = {"海底捞评论.xlsx":"content"}
    for p in my_dict:
        data = pd.read_excel(p)
        for i in range(len(data.loc[:, "content"])):
            print("%s" % p + "正在进行中，进度%s/%s" % (str(i + 1), str(len(data.loc[:, "content"]))))
            sentence = data.loc[i, "content"]
            my_word_list = sent2word(sentence)
            my_dict = sentiment_loc_with_score(my_word_list)
            my_score, my_length = count_score(my_dict)
            try:
               data.loc[i, "score"] = my_score / float(my_length)
            except:
               data.loc[i, "score"] = "没有情感词"
            data.loc[i, "length"] = my_length

        data.to_excel("result/"+p)