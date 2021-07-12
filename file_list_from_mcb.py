import docx
import json
import os
import re

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
def mcb_filter(mcb_list):
    #Remove Null type items
    mcb_list = list(filter(None, mcb_list))
    in_list = []
    out_list = []
    for mcb in mcb_list:
        if re.match(r"BN\d+", str(mcb)):
            in_list.append(mcb)
        else:
            out_list.append(mcb)
    return in_list, out_list
def get_mcb_list(json_file_path):
    with open (r"HCM_patients_woLink.json") as f:
        mcb_list = json.load(f)
    mcb_list, _ = mcb_filter(mcb_list)
    return mcb_list

def get_file_with_mcb(directory_path, mcb_file_path):
    file_list = get_all_document(directory_path)
    mcb_list = get_mcb_list(mcb_file_path)

    files_with_mcb = set()

    for f in file_list:
        document_string = docx_to_string(f)
        for mcb in mcb_list:

            if mcb[2:].strip() in document_string:
                files_with_mcb.add(f)
    return list(files_with_mcb)

def files_with_mcb_to_txt(mcb_file_list, out_file_path = "out.txt"):
    with open(out_file_path, 'w', encoding='utf-8') as f:
        for file_name in mcb_file_list:
            f.write(file_name)
            f.write("\n")
if __name__ == "__main__":
