import docx
import os
import re
import pandas as pd

def getPersonalInformation(document):
    lst=[]
    for paragraph in document.paragraphs:
        lst.append(paragraph.text)
    
    text = '\n'.join(lst)
    start = text.find("Thông tin ca bệnh")
    end = text.find("Lịch sử đi lại")

    dict = getPersonalInformationDetail(text[start:end])
    return cleanData(dict)

def cleanData(dict):
    cleanedDict = {}
    cleanedDict["hoTen"] = dict["hoTen"][0][0]
    cleanedDict["maBN"] = dict["maBN"][0]
    cleanedDict["namSinh"] = dict["namSinh"][0][-4:]
    cleanedDict["gioiTinh"] = dict["gioiTinh"][0][1:][0]

    if len(dict["CMND"][0][0]) > 8:
        cleanedDict["CMND"] = re.findall('[0-9]+', dict["CMND"][0][0])[0]
    else: 
        cleanedDict["CMND"] =""

    startIndex = dict["quocTich"][0][0].find(":")
    cleanedDict["quocTich"] = dict["quocTich"][0][0][startIndex+1:]
    cleanedDict["SDT"] = dict["SDT"][-1]
    return cleanedDict


def getPersonalInformationDetail(text):
    dict = {
        "hoTen" : "(([A-ZẮẰẲẴẶĂẤẦẨẪẬÂÁÀÃẢẠĐẾỀỂỄỆÊÉÈẺẼẸÍÌỈĨỊỐỒỔỖỘÔỚỜỞỠỢƠÓÒÕỎỌỨỪỬỮỰƯÚÙỦŨỤÝỲỶỸỴ']+\s?){2,})",
        "maBN":"(BN\s?[0-9]+)",
        "namSinh":"(sinh năm:? \d{4})",
        "gioiTinh":"(\s?(nam|nữ|NAM|NỮ|Nam|Nữ))",
        "CMND":"(((nhân dân)|(CCCD)):\s?\d{8,})",
        "quocTich": "(tịch: ([a-zắằẳẵặăấầẩẫậâáàãảạđếềểễệêéèẻẽẹíìỉĩịốồổỗộôớờởỡợơóòõỏọứừửữựưúùủũụýỳỷỹỵA-ZẮẰẲẴẶĂẤẦẨẪẬÂÁÀÃẢẠĐẾỀỂỄỆÊÉÈẺẼẸÍÌỈĨỊỐỒỔỖỘÔỚỜỞỠỢƠÓÒÕỎỌỨỪỬỮỰƯÚÙỦŨỤÝỲỶỸỴ\s])+)",
        "SDT":"(\d{4}.?\d{3}.?\d{3})"
    }

    for regex in dict:
        if re.search(re.compile(dict[regex]),text):
            dict[regex] = [m for m in re.findall(re.compile(dict[regex]),text)]
    return dict


if __name__ == '__main__':
    #path docx file
    file_path = r'C:\Users\TRANCONGMINH\Covid19-IE\BÁO CÁO FILE WORD\BC CHUỖI HỘI THÁNH - NHÓM 4\BN6298_LÊ THIÊN ÂN HỒNG.docx'
    document = docx.Document(file_path)
    dict = getPersonalInformation(document)
    print(dict)




