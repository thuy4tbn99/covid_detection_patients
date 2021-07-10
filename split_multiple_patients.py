import re
import docx
def docx_to_string(docx_file):
    try:
        document = docx.Document(docx_file)
    except:
        print("Unable to read file" + docx_file)
        return ""
    return "\n".join([paragraph.text for paragraph in document.paragraphs])
def remove_last_line(s):
    return s[:s.rfind('\n')]
def split_normal_multiple(document_path):
    """Return a list contains infomation for each patient in the word document of normal type"""
    document_string = docx_to_string(document_path)
    document_string_lower = document_string.lower()
    anchor = "thông tin ca bệnh"
    pos = [m.start() for m in re.finditer(anchor, document_string_lower)]
    pos.append(len(document_string_lower))
    i = 0
    splitted_patient_info = []
    while i < len(pos) - 1:
        splitted_patient_info.append(document_string[pos[i]:pos[i+1]])
        i += 1

    i = 0
    while i < len(pos) -2:
        splitted_patient_info[i]= remove_last_line(splitted_patient_info[i].strip())
        i+=1
    return splitted_patient_info

if __name__ == "__main__":
    for BN in split_normal_multiple(r"baocao_covid\ BC CHUỖI VỰA VE CHAI\BN12400_BN12491_BN13209_BN12896_ĐOÀN THANH VÂN_NGUYỄN THỊ THÙY NGA_ĐOÀN GIA HUY_ĐOÀN THỊ NHƯ NGỌC_ NHÓM 04.docx"):
        print(BN)
        print("-"*100)