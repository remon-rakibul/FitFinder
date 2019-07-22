import os
import re
import heapq
import convert2txt
from nltk import sent_tokenize
from nltk import word_tokenize
from nltk.corpus import stopwords as stwrd
from wordcloud import WordCloud, STOPWORDS
    
def preprocess(data):
    sentences = sent_tokenize(data)
    for i in range(len(sentences)):
        words = word_tokenize(sentences[i])
        words = [word for word in words if word not in stwrd.words('english')]
        sentences[i] = ' '.join(words)
    for i in range(len(sentences)):
            sentences[i] = re.sub(r'\\t+|\\r+', ' ', sentences[i])
            sentences[i] = re.sub(r'â+', ' ', sentences[i])
            sentences[i] = re.sub(r'ï+', ' ', sentences[i])
            sentences[i] = re.sub(r'»+', ' ', sentences[i])
            sentences[i] = re.sub(r'¿Â+', ' ', sentences[i])
            sentences[i] = re.sub(r'Â', ' ', sentences[i])
            sentences[i] = re.sub(r'§+', ' ', sentences[i])
            sentences[i] = re.sub(r'Ã+', ' ', sentences[i])
            sentences[i] = re.sub(r'\W+', ' ', sentences[i])
            sentences[i] = re.sub(r'\s+', ' ', sentences[i])
    return sentences

def freq_words(sentences):
    word2count = {}
    
    for sentence in sentences:
        words = word_tokenize(sentence)
        for word in words:
            if word not in word2count.keys():
                word2count[word] = 1
            else:
                word2count[word] += 1
    
    freq_words = heapq.nlargest(200, word2count, key=word2count.get)
    
    text = ' '.join(freq_words)
    return text

def create_wordcloud(text):
    stopwords = set(STOPWORDS)
    
    wordcloud = WordCloud(
            width=1600, height=1200, background_color='white',
            stopwords=stopwords, min_font_size=10).generate(text)
    
#    wordcloud.to_file(filename'.png')
    return wordcloud

def write_wordcloud(wc_dir, filename, wordcloud):
    try:
        os.mkdir(wc_dir)
    except:
        pass
    current = os.getcwd()
    os.chdir(wc_dir)
    if filename.endswith('.docx'):
        wordcloud.to_file(filename[:-5]+'.png')
    else:
        wordcloud.to_file(filename[:-5]+'.png')
    os.chdir(current)

def text_to_cloud(data, filename, wc_dir):
    sentences = preprocess(data)
    text = freq_words(sentences)
    wordcloud = create_wordcloud(text)
    write_wordcloud(wc_dir, filename, wordcloud)
    
def wordcloud_a_dir(cv_dir, wc_dir):
    files = os.listdir(cv_dir)
    for file in files:
        if file.endswith('.pdf'):
            data = convert2txt.extract_text(cv_dir+file, '.pdf')
            text_to_cloud(data, file, wc_dir)
        elif file.endswith('.doc'):
            data = convert2txt.extract_text(cv_dir+file, '.doc')
            text_to_cloud(data, file, wc_dir)
        elif file.endswith('.docx'):
            data = convert2txt.extract_text(cv_dir+file, '.docx')
            text_to_cloud(data, file, wc_dir)
        elif file.endswith('.txt'):
            with open(cv_dir+file, encoding='utf-8') as f:
                data = f.read().replace('\n', '')
            text_to_cloud(data, file, wc_dir)
            
wordcloud_a_dir('cv/txt/', 'cv_wordclouds')