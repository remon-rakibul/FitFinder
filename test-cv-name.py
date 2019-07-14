import os
import re
import csv
import spacy
import preprocess
import convert2txt
from nltk.tokenize import sent_tokenize

nlp = spacy.load('name-cv-bn')

def extract_names(cv_dir, word_limit):
    extracted_names = {}
    files = os.listdir(cv_dir)
    for file in files:
        if file.endswith('.pdf'):
            text = convert2txt.extract_text(cv_dir+file, '.pdf')
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            extracted_names[file] = []
            for e in nlp_text.ents:
                extracted_names[file].append(e.text)
        elif file.endswith('.doc'):
            text = convert2txt.extract_text(cv_dir+file, '.doc')
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            extracted_names[file] = []
            for e in nlp_text.ents:
                extracted_names[file].append(e.text)
        elif file.endswith('.docx'):
            text = convert2txt.extract_text(cv_dir+file, '.docx')
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            extracted_names[file] = []
            for e in nlp_text.ents:
                extracted_names[file].append(e.text)
        elif file.endswith('.txt'):
            with open(cv_dir+file, encoding='utf-8') as f:
                text = f.read()
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            extracted_names[file] = []
            for e in nlp_text.ents:
                extracted_names[file].append(e.text)
    return extracted_names
'''
###################### test on each sentence ######################
def extract_names_each_sent(cv_dir, word_limit, sent_limit):
    extracted_names = {}
    files = os.listdir(cv_dir)
    for file in files:
        if file.endswith('.pdf'):
            text = convert2txt.extract_text(cv_dir+file, '.pdf')
            extracted_names[file] = []
            words = text.split()
            text = ' '.join(words[:word_limit])
            sentences = sent_tokenize(text)
            sentences = sentences[:sent_limit]
            for sentence in sentences:
                text = preprocess.process(sentence)
                nlp_text = nlp(text)
                for e in nlp_text.ents:
                    extracted_names[file].append(e.text)
        elif file.endswith('.doc'):
            text = convert2txt.extract_text(cv_dir+file, '.doc')
            extracted_names[file] = []
            words = text.split()
            text = ' '.join(words[:word_limit])
            sentences = sent_tokenize(text)
            sentences = sentences[:sent_limit]
            for sentence in sentences:
                text = preprocess.process(sentence)
                nlp_text = nlp(text)
                for e in nlp_text.ents:
                    extracted_names[file].append(e.text)
        elif file.endswith('.docx'):
            text = convert2txt.extract_text(cv_dir+file, '.docx')
            extracted_names[file] = []
            words = text.split()
            text = ' '.join(words[:word_limit])
            sentences = sent_tokenize(text)
            sentences = sentences[:sent_limit]
            for sentence in sentences:
                text = preprocess.process(sentence)
                nlp_text = nlp(text)
                for e in nlp_text.ents:
                    extracted_names[file].append(e.text)
        elif file.endswith('.txt'):
            with open(cv_dir+file, encoding='utf-8') as f:
                text = f.read()                
            extracted_names[file] = []
            words = text.split()
            text = ' '.join(words[:word_limit])
            sentences = sent_tokenize(text)
            sentences = sentences[:sent_limit]
            for sentence in sentences:
                text = preprocess.process(sentence)
                nlp_text = nlp(text)
                for e in nlp_text.ents:
                    extracted_names[file].append(e.text)
    return extracted_names
'''

def export_to_csv(csv_name, dic):
    with open(csv_name, 'w', newline='') as csv_file:
        headers = ['CV Name', 'Extracted Name', 'Detection', 'Total Failed']
        writer = csv.DictWriter(csv_file, fieldnames=headers)
        writer.writeheader()
        failed = 0
        for file_name, extracted_name in dic.items():
            if len(extracted_name) == 0:
                detection = 0
                failed += 1
            else:
                detection = 1
            writer.writerow(
                        {
                                headers[0]: file_name,
                                headers[1]: ' '.join(extracted_name),
                                headers[2]: detection
                        }
                    )
        writer.writerow({headers[3]: failed})

def extract_emails(cv_dir, word_limit):
    extracted_emails = {}
    files = os.listdir(cv_dir)
    pattern = re.compile(r'[\w\.-]+@[\w\.-]+')
    for file in files:
        if file.endswith('.pdf'):
            text = convert2txt.extract_text(cv_dir+file, '.pdf')
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_emails[file] = results
        elif file.endswith('.doc'):
            text = convert2txt.extract_text(cv_dir+file, '.doc')
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_emails[file] = results
        elif file.endswith('.docx'):
            text = convert2txt.extract_text(cv_dir+file, '.docx')
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_emails[file] = results
        elif file.endswith('.txt'):
            with open(cv_dir+file, encoding='utf-8') as f:
                text = f.read()
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_emails[file] = results
    return extracted_emails

def extract_mobile_numbers(cv_dir, word_limit):
    extracted_mobile_numbers = {}
    files = os.listdir(cv_dir)
    pattern = re.compile(r"\(?\+?[8]{2}?0?\)?\0?-?0?[0-9]{3}-?[0-9]{3}-?[0-9]{4}|[0-9]{4}-?[0-9]{3}-?[0-9]{4}|[0-9]{5}-[0-9]{6}")
    for file in files:
        if file.endswith('.pdf'):
            text = convert2txt.extract_text(cv_dir+file, '.pdf')
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_mobile_numbers[file] = results
        elif file.endswith('.doc'):
            text = convert2txt.extract_text(cv_dir+file, '.doc')
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_mobile_numbers[file] = results
        elif file.endswith('.docx'):
            text = convert2txt.extract_text(cv_dir+file, '.docx')
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_mobile_numbers[file] = results
        elif file.endswith('.txt'):
            with open(cv_dir+file, encoding='utf-8') as f:
                text = f.read()
            words = text.split()
            text = ' '.join(words[:word_limit])
            results = pattern.findall(text)
            extracted_mobile_numbers[file] = results
    return extracted_mobile_numbers
    

extracted_names = extract_names('cv/txt/', 25)

#extracted_names_each_sent = extract_names_each_sent('cv/txt/', 25, 3)

extracted_emails = extract_emails('cv/txt/', 150)

extracted_mobile_numbers = extract_mobile_numbers('cv/txt/', 100)

export_to_csv('report.csv', extracted_names)

#export_to_csv('report_each_sent.csv', extracted_names_each_sent)