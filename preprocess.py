import re
import spacy
nlp = spacy.load('en_core_web_sm')

def process(txt):
    data = txt
    data = re.sub(r'\W+', ' ', data)
    data = data.lower()
    data = re.sub(r'\\t+|\\r+', ' ', data)
    data = re.sub(r'â+', ' ', data)
    data = re.sub(r'ï+', ' ', data)
    data = re.sub(r'\d+', ' ', data)
    data = re.sub(r'\s+', ' ', data)
    doc = nlp(u'{}'.format(data))
    tokens = [token.text for token in doc if not token.is_stop]
    data = ' '.join(tokens)
    return data