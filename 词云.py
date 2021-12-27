import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# 导入文本数据并进行简单的文本处理
# 去掉换行符和空格
text = open("C:/Users/xiaobin.ma/Downloads/工作簿1.txt",encoding='utf8').read()
text = text.replace('\n',"").replace("\u3000","")

# 分词，返回结果为词的列表
text_cut = jieba.lcut(text)
# 将分好的词用某个符号分割开连成字符串
text_cut = ' '.join(text_cut)

stop_words = open("C:/Users/xiaobin.ma/Downloads/stop_word.txt",encoding="utf8").read().split("\n")

word_cloud = WordCloud(font_path="simsun.ttc",  # 设置词云字体
                       background_color="white",# 词云图的背景颜色
                       stopwords=stop_words,    # 去掉的停词
                       colormap = 'PuBu') 
word_cloud.generate(text_cut)

# 运用matplotlib展现结果
plt.subplots(figsize=(12,8))
plt.imshow(word_cloud)
plt.axis("off")