import docx
import re
import os
from enum import Enum

class document_type(Enum):
    NORMAL_SINGLE = 1
    NORMAL_MULTIPLE = 2
    QUICK_REPORT = 3
    QUICK_REPORT_2 = 4
    OTHERS = 5

def docx_to_string(docx_file):
    try:
        document = docx.Document(docx_file)
    except:
        print("Unable to read file" + docx_file)
        return ""
    return "\n".join([paragraph.text for paragraph in document.paragraphs])

def get_all_document(directory_path):
    """Return a list contains all docx file path in a given directory"""
    file_paths = []
    for subdir, dirs, files in os.walk(directory_path):
        for file_name in files:
            file_path = subdir + os.sep + file_name
            if file_path.endswith(".docx"):
                file_paths.append(file_path)
    return file_paths

class DocumentClassifier:
    def __init__(self):
        pass
        
    def check_document_type(self, document_path):
        _, file_name = os.path.split(document_path)
        BN_count_filename = len(re.findall(r'[B][N,n]\d{1,6}|[B][N][_]', file_name))
        
        patient_count_filename = 0
        i = 0
        for name in  re.finditer(r"(?<=_)((([A-ZẮẰẲẴẶĂẤẦẨẪẬÂÁÀÃẢẠĐẾỀỂỄỆÊÉÈẺẼẸÍÌỈĨỊỐỒỔỖỘÔỚỜỞỠỢƠÓÒÕỎỌỨỪỬỮỰƯÚÙỦŨỤÝỲỶỸỴ']+\s?){3,5}))(?=_)", file_name):
            if ' ' in name.group().split():
                patient_count_filename+=1
        document_string = docx_to_string(document_path).lower()
        #only normal type files contain this string
        n = document_string.find("báo cáo nhanh thông tin về")
        pattern = re.compile(r"(?<=báo cáo nhanh thông tin về) +\d+ +((?=trường hợp)|(?=bệnh nhân))")
        if n >= 0:
            patient_count = document_string.count("thông tin ca bệnh")
            found_patient_num = re.search(pattern, document_string)
            if found_patient_num:
                patient_num_string = found_patient_num.group().strip() 
                try:
                    patient_num = int(patient_num_string)
                except:
                    return document_type.OTHERS
                # if int(patient_count == patient_num) + int(patient_count_in_file_name == patient_count) +int(patient_num ==patient_count_in_file_name) < 2:
                if patient_count != patient_num or BN_count_filename != patient_count:
                    return document_type.OTHERS
                if patient_num >= 2 or patient_count_filename >= 2:
                    return document_type.NORMAL_MULTIPLE
                return document_type.NORMAL_SINGLE
            return document_type.OTHERS
        #only quick report type files contains this tring
        n = document_string.find("báo cáo nhanh")
        if n >= 0:
            f = document_string.find("thông tin người f")
            if f >= 0: 
                return document_type.QUICK_REPORT
            return document_type.QUICK_REPORT_2
        return document_type.OTHERS

    def categorize(self, directory_path):
        doc_classes = {
            "normal_single": [],
            "normal_multiple":[],
            "quick_report": [],
            "quick_report2":[],
            "others": []
        }
        documents_path = get_all_document(directory_path)
        for path in documents_path:
            t = self.check_document_type(path)
            if t == document_type.NORMAL_SINGLE:
                doc_classes['normal_single'].append(path)
            elif t == document_type.NORMAL_MULTIPLE:
                doc_classes['normal_multiple'].append(path)
            elif t == document_type.QUICK_REPORT:
                doc_classes['quick_report'].append(path)
            elif t == document_type.QUICK_REPORT_2:
                doc_classes['quick_report2'].append(path)
            else:
                doc_classes['others'].append(path)
        
        doc_class_size = {}
        total = 0
        for key in doc_classes:
            doc_class_size[key] = len(doc_classes[key])
            total += len(doc_classes[key])
        doc_class_size['total'] = total
        return doc_classes, doc_class_size