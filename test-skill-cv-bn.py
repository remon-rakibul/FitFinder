import os
import re
import csv
import spacy
import convert2txt

nlp = spacy.load('skill-cv-bn')

def create_dic(filename):
    result = {
                'Filename': filename,
                'Name':[],
                'Number':[],
                'Email':[],
                'Experience':[],
                'Companies':[],
                'Designation':[],
                'Responsibilities':[],
                'Skills':[],
                'Certification':[],
                'Degree':[],
                'Institution':[],
                'Edu_subject':[],
                'Graduation_year':[]
                }
    return result      

def get_info(text, dic_name):
    nlp_text = nlp(text)
    for e in nlp_text.ents:
        if e.label_ == 'Name':
            dic_name['Name'].append(e.text)
        elif e.label_ == 'Number':
            dic_name['Number'].append(e.text)
        elif e.label_ == 'Email':
            dic_name['Email'].append(e.text)
        elif e.label_ == 'Experience':
            dic_name['Experience'].append(e.text)
        elif e.label_ == 'Companies':
            dic_name['Companies'].append(e.text)
        elif e.label_ == 'Designation':
            dic_name['Designation'].append(e.text)
        elif e.label_ == 'Responsibilities':
            dic_name['Responsibilities'].append(e.text)
        elif e.label_ == 'Skills':
            dic_name['Skills'].append(e.text)
        elif e.label_ == 'Certification':
            dic_name['Certification'].append(e.text)
        elif e.label_ == 'Degree':
            dic_name['Degree'].append(e.text)
        elif e.label_ == 'Institution':
            dic_name['Institution'].append(e.text)
        elif e.label_ == 'Edu_subject':
            dic_name['Edu_subject'].append(e.text)
        elif e.label_ == 'Graduation_year':
            dic_name['Graduation_year'].append(e.text)

def extract_info(cv_dir):
    extracted_info = []
    files = os.listdir(cv_dir)
    for file in files:
        if file.endswith('.pdf'):
            text = convert2txt.extract_text(cv_dir+file, '.pdf')
            result = create_dic(file)
            get_info(text, result)
            extracted_info.append(result)
        elif file.endswith('.doc'):
            text = convert2txt.extract_text(cv_dir+file, '.doc')
            result = create_dic(file)
            get_info(text, result)
            extracted_info.append(result)
        elif file.endswith('.docx'):
            text = convert2txt.extract_text(cv_dir+file, '.docx')
            result = create_dic(file)
            get_info(text, result)
            extracted_info.append(result)
        elif file.endswith('.txt'):
            with open(cv_dir+file, encoding='utf-8-sig') as f:
                text = f.read()
            result = create_dic(file)
            get_info(text, result)
            extracted_info.append(result)
    return extracted_info

def process(txt):
    data = txt
    data = re.sub(r'â+', ' ', data)
    data = re.sub(r'ï+', ' ', data)
    data = re.sub(r'\\u+', ' ', data)
    data = re.sub(r'â€+', ' ', data)
    data = re.sub(r'™+', ' ', data)
    data = re.sub(r'â€+', ' ', data)
    data = re.sub(r'â+', ' ', data)
    data = re.sub(r'\s+', ' ', data)
    return data

def create_csv_row_dic(headers, dic):
    result = {}
    for i in range(len(headers)):
        if i == 0:
            result[headers[i]] = dic[headers[i]]
        else:
            result[headers[i]] = ', '.join([process(i) for i in dic[headers[i]]])
    return result

def export_to_csv(csv_name, extracted_info):
    with open(csv_name, 'w', newline='', encoding="utf-8") as csv_file:
        headers = [
                'Filename',
                'Name',
                'Number',
                'Email',
                'Experience',
                'Companies',
                'Designation',
                'Responsibilities',
                'Skills',
                'Certification',
                'Degree',
                'Institution',
                'Edu_subject',
                'Graduation_year'
                ]
        writer = csv.DictWriter(csv_file, fieldnames=headers)
        writer.writeheader()
        for dic in extracted_info:
            write_dic = create_csv_row_dic(headers, dic)
            writer.writerow(write_dic)

export_to_csv('skills_report.csv', extract_info('../cv/txt/'))