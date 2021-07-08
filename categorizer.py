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

def check_document_type(document_path):
    document_string = docx_to_string(document_path).lower()
    #only normal type files contain this string
    n = document_string.find("báo cáo nhanh thông tin về")
    pattern = re.compile(r"(?<=báo cáo nhanh thông tin về) +\d+ +((?=trường hợp)|(?=bệnh nhân))")
    if n >= 0:
        patient_num_search = re.search(pattern, document_string)
        if patient_num_search:
            patient_num_string = patient_num_search.group().strip()
            try:
                patient_num = int(patient_num_string)
            except:
                return document_type.OTHERS
            if patient_num >= 2:
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

def categorize(directory_path):
    normal_single = []
    normal_multiple = []
    quick_report = []
    quick_report2 = []
    others = []
    documents_path = get_all_document(directory_path)
    for path in documents_path:
        t = check_document_type(path)
        if t == document_type.NORMAL_SINGLE:
            normal_single.append(path)
        elif t == document_type.NORMAL_MULTIPLE:
            normal_multiple.append(path)
        elif t == document_type.QUICK_REPORT:
            quick_report.append(path)
        elif t == document_type.QUICK_REPORT_2:
            quick_report2.append(path)
        else:
            others.append(path)
    return {
        "normal_single":normal_single,
        "normal_multiple":normal_multiple,
        "quick_report": quick_report,
        "quick_report2":quick_report2,
        "others":others
    }
if __name__ == "__main__":
    "do nothing"