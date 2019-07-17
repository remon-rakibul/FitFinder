import os
import re
import csv
import spacy
import preprocess
import convert2txt
from nltk.tokenize import sent_tokenize

nlp = spacy.load('name-cv-bn')

def extract_mobile_numbers(text):
    pattern = re.compile(r"\(?\+?[8]{2}?0?\)?\0?-?0?[0-9]{3}-?[0-9]{3}-?[0-9]{4}|[0-9]{4}-?[0-9]{3}-?[0-9]{4}|[0-9]{5}-[0-9]{6}")
    results = pattern.findall(text)
    return results

def extract_emails(text):
    pattern = re.compile(r'[\w\.-]+@[\w\.-]+')
    results = pattern.findall(text)
    return results

def extract_info(cv_dir, word_limit):
    extracted_info = []
    files = os.listdir(cv_dir)
    for file in files:
        if file.endswith('.pdf'):
            text = convert2txt.extract_text(cv_dir+file, '.pdf')
            emails = extract_emails(text)
            numbers = extract_mobile_numbers(text)
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            name = []
            for e in nlp_text.ents:
                name.append(e.text)
            extracted_info.append({
                    'filename': file,
                    'name': ' '.join(name),
                    'number': numbers,
                    'email': emails
                    })
        elif file.endswith('.doc'):
            text = convert2txt.extract_text(cv_dir+file, '.doc')
            emails = extract_emails(text)
            numbers = extract_mobile_numbers(text)
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            name = []
            for e in nlp_text.ents:
                name.append(e.text)
            extracted_info.append({
                    'filename': file,
                    'name': ' '.join(name),
                    'number': numbers,
                    'email': emails
                    })
        elif file.endswith('.docx'):
            text = convert2txt.extract_text(cv_dir+file, '.docx')
            emails = extract_emails(text)
            numbers = extract_mobile_numbers(text)
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            name = []
            for e in nlp_text.ents:
                name.append(e.text)
            extracted_info.append({
                    'filename': file,
                    'name': ' '.join(name),
                    'number': numbers,
                    'email': emails
                    })
        elif file.endswith('.txt'):
            with open(cv_dir+file, encoding='utf-8') as f:
                text = f.read()
            emails = extract_emails(text)
            numbers = extract_mobile_numbers(text)
            words = text.split()
            text = ' '.join(words[:word_limit])
            text = preprocess.process(text)
            nlp_text = nlp(text)
            name = []
            for e in nlp_text.ents:
                name.append(e.text)
            extracted_info.append({
                    'filename': file,
                    'name': ' '.join(name),
                    'number': numbers,
                    'email': emails
                    })
    return extracted_info

def export_to_csv(csv_name, extracted_info):
    with open(csv_name, 'w', newline='') as csv_file:
        headers = ['CV Name', 'Extracted Name', 'Detection', 'Numbers', 'Emails', 'Total Failed']
        writer = csv.DictWriter(csv_file, fieldnames=headers)
        writer.writeheader()
        failed = 0
        for dic in extracted_info:
            if len(dic['name']) == 0:
                detection = 0
                failed += 1
            else:
                detection = 1
            writer.writerow(
                        {
                                headers[0]: dic['filename'],
                                headers[1]: dic['name'],
                                headers[2]: detection,
                                headers[3]: dic['number'],
                                headers[4]: dic['email']
                        }
                    )
        writer.writerow({headers[5]: failed})



extracted_info = extract_info('cv/txt/', 25)

export_to_csv('report_new.csv', extracted_info)