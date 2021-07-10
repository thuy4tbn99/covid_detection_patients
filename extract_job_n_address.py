import docx
import re
def docx_to_string(docx_file):
    try:
        document = docx.Document(docx_file)
    except:
        print("Unable to read file")
        return ""
    return "\n".join([paragraph.text for paragraph in document.paragraphs])

def extract_sections(document_string, section):
   """"Return the string which contains the content of each section"""
   begin = ""
   end = ""
   if section == 1:
      begin = document_string.find("Thông tin ca bệnh")
      end = document_string.find("Lịch sử đi lại và tiền sử tiếp xúc và triệu chứng lâm sàng")
   elif section == 2:
      begin = document_string.find("Lịch sử đi lại và tiền sử tiếp xúc và triệu chứng lâm sàng")
      end = document_string.find("Các hoạt động đã triển khai")
   elif section == 3:
      begin = document_string.find("Các hoạt động đã triển khai")
      end = len(document_string)
   section_string = document_string[begin:end]
   return section_string

def find_job(section_string):
    """Return patient's job description from patient's info string"""
    pattern = re.compile(r'(?<=Nghề nghiệp:)[^\n]*(?=\n)|(?<=Tên và địa chỉ nơi làm việc:)[^\n]*(?=\n)|(?<=Tên và địa chỉ làm việc)[^\n]*(?=\n)|(?<=Địa chỉ nơi ở và nơi làm)[^\n]*(?=\n)')
    job_search = re.search(pattern, section_string)
    if job_search:
        return job_search.group()

def find_address(document_string):
    """Return patient's address from patient's info stirng"""
    address = ""
    pattern = re.compile(r'(?<=Địa chỉ:)[^\n]+|(?<=Địa chỉ nơi ở:)[^\n]+|(?<=Địa chỉ tạm trú:)[^\n]+|(?<=Địa chỉ nơi ở hiện nay:)[^\n]+|(?<=Địa chỉ nơi)[^\n]+|(?<=Địa chỉ nơi ở và nơi làm)[^\n]+|(?<=Địa chỉ nhà:)[^\n]+')
    if re.search(re.compile(pattern),document_string):
        address = re.findall(pattern, document_string)[0]
    return address

def split_address_normal(address_string):
        address_string = address_string.replace('.', '')
        districts = [ 'quận 11','q11', 'quận 12','q12' ,'quận 10','q10','quận 9','q9', 'quận 4','q4' ,   'quận 6','q6',
            'quận 2','q2','quận 5', 'q5', 'quận 7', 'q7','q3' ,'quận 3',  'quận 1', 'q1' ,
            'cần giờ','củ chi','gò vấp', 'phú nhuận',  'bình thạnh', 'quận 8','q8', 'tân bình', 'nhà bè',  
            'hóc môn', 'bình chánh','thủ đức', 'tân phú','bình tân']
        villages = ['phường 16', 'p16', 'tăng nhơn phú b', 'thủ thiêm', 'linh đông', 'bình trưng tây', 'hoà thạnh', 
            'phường 08','p8', 'thảo điền', 'tân chánh hiệp', 'hiệp thành', 'thạnh an', 'vĩnh lộc a', 
            'bình thuận', 'phú xuân', 'tân an hội', 'nhơn đức', 'phường 11','p11' ,'long thới', 'hóc môn', 
            'an nhơn tây', 'phước long b', 'trường thọ', 'bình trưng đông', 'an phú đông', 'bình chiểu', 'bình mỹ',
            'cầu kho', 'thái mỹ', 'phạm văn hai', 'hưng long', 'tân hiệp', 'thạnh xuân', 'trung mỹ tây', 'hiệp bình chánh', 
            'phú thuận', 'tân hưng', 'an lạc', 'tân thới hoà', 'thạnh lộc', 'long thạnh mỹ', 'bình hưng hoà b', 'bà điểm', 
            'tân quý tây', 'bình khánh', 'phường 28', 'p28', 'hiệp phước', 'phú trung', 'bình hưng hòa', 
            'đông hưng thuận', 'phường 7','p7', 'nhị bình', 'tân kiểng', 'an lợi đông', 'linh chiểu', 'an khánh', 
            'phú hòa đông', 'phạm ngũ lão', 'bình an', 'tân thạnh tây', 'phước long a', 'bình chánh', 'linh trung', 'bình trị đông', 'bến nghé', 'long bình', 'tân thuận tây', 'tăng nhơn phú a', 'phước thạnh', 
            'tam phú', 'xuân thới thượng', 'phường 12','p12', 'trung lập thượng', 'tân sơn nhì', 'thới tam thôn', 
            'tân tạo a', 'bình trị đông a', 'phường 24','p24', 'phước bình', 'bình trị đông b', 'phước lộc', 
            'tân qúy', 'đông thạnh', 'lê minh xuân', 'tân nhựt', 'tam bình', 'phú hữu', 'phước kiển', 'tây thạnh', 
            'an thới đông', 'phường 10', 'p10', 'đa kao', 'nhà bè', 'hiệp tân', 'phạm văn cội', 'phú thọ hoà', 
            'quy đức', 'an phú tây', 'tân xuân', 'phú mỹ hưng', 'phường 27', 'p27', 'linh xuân', 'phường 26', 'p26', 
            'nguyễn thái bình', 'phước vĩnh an', 'linh tây', 'tân phong', 'bến thành', 'hiệp phú', 
            'bình hưng hoà a', 'phường 21', 'long phước', 'trung lập hạ', 'phường 17', 'p17', 'tân thạnh đông', 
            'cô giang', 'xuân thới đông', 'bình lợi', 'an phú', 'đa phước', 'phú mỹ', 'tân hưng thuận', 'bình thọ', 
            'phường 18','p18', 'tân tạo', 'phước hiệp', 'cần thạnh', 'phường 13', 'p13','tân quy', 'hiệp bình phước', 
            'phường 15', 'phường 05', 'long hòa', 'hòa phú', 'cầu ông lãnh', 'cát lái', 'thới an', 'tân thuận đông', 
            'tân thới hiệp', 'xuân thới sơn', 'thạnh mỹ lợi', 'tân kiên', 'lý nhơn', 'an lạc a', 'trung an', 'phong phú', 
            'bình hưng', 'nguyễn cư trinh', 'phường 3', 'p3', 'phường 6','p6' ,'phường 19','p19','tân túc', 'phú thạnh', 'phường 14','p14', 'tân phú trung', 'tân phú', 'tân thới nhì', 'phường 22', 'p22',
            'tân thông hội', 'sơn kỳ', 'trung chánh', 'tân định', 'tân thới nhất', 'tam thôn hiệp', 'vĩnh lộc b', 
            'phường 25','p25', 'phường 9','p9', 'củ chi', 'tân thành', 'nhuận đức', 'long trường', 'trường thạnh', 
            'phường 1', 'p1', 'phường 2', 'p2', 'phường 4','p4']
        hcm = {'tp hcm', 'tphcm', 'tp hồ chí minh', 'thành phố hồ chí minh'}
        street = ""
        village = ""
        district = ""
        provine = ""
        for i in hcm:
                if i in address_string:
                    provine = "TP HCM"
                    address_string = address_string[:address_string.find(i)]
                    break
        for d in districts:
                if d in address_string:
                    if d == "tân phú":
                        break
                    district = d
                    provine = "TP HCM"
                    address_string = address_string[:address_string.find(d)]
                    break
        for v in villages:
            if v in address_string:
                if v == "tân phú":
                    break
                village = v
                provine = 'TP HCM'
                street= address_string[:address_string.find(v)]
                street = street.replace("xã", "")
                street = street.replace("phường", "")
                street = street.replace(",", "")
                break
        return street, village, district, provine

def extract_patient_info(document_string):
    """"Return a json file which contains patient's information extracted from a word document"""
    # document_string = docx_to_string(docx_file)
    patient_info_section = extract_sections(document_string, 1)
    #get patient's job's description
    patient_job = find_job(patient_info_section)
    #get patient's address location
    patient_address = find_address(patient_info_section).lower()
    street, village, district, provine = split_address_normal(patient_address)
    output = {
        # "file_name": docx_file,
        "nghe_nghiep":patient_job,
        "dia chi": patient_address,
        "duong/thon/xom": street,
        "xa/phuong": village,
        "quan/huyen": district,
        "tinh/thanhpho": provine
    }
    return output
if __name__ == "__main__":
    "do nothing"
