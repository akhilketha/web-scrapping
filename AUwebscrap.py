import warnings
import requests
import contextlib
from bs4 import BeautifulSoup
from openpyxl import Workbook
import urllib3

#Generates col names:
def generate_col_names(number):
    try:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        res = requests.post('https://aucoe.info/RDA/resultsnew/result_grade.php', data={'serialno': 1048,
                                                                                        'course': 'B.E./B.TECH/INTEGRATED COURSE THIRD YEAR SECOND SEMESTER',
                                                                                        'degree': 'B.E./B.TECH/Integrated Course',
                                                                                        'table': 'gradestructure3',
                                                                                        'appearing_year': 'APRIL 2018',
                                                                                        'Date_time': '2018-07-06 11:39:44',
                                                                                        'regno': number,
                                                                                        'revdate': '2018-07-19',
                                                                                        'revfee': 750
                                                                                        },verify=False)

        soup = BeautifulSoup(res.text,"lxml")
        col_names=[]
        for table in soup.find_all("table")[3:4]:
            trs = table.find_all("tr")
            col_names.append(trs[0].text.split(':')[0].strip())
            col_names.append(trs[1].text.split(':')[0].strip())

        for table in soup.find_all("table")[4:5]:
            trs = table.find_all("tr")[1:]
            for tr in trs:
                tds = tr.find_all('td')
                col_names.append(tds[0].text)
        return col_names
    except Exception as e:
        pass

def generate_individual_marks(number):
    try:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        res = requests.post('https://aucoe.info/RDA/resultsnew/result_grade.php', data={'serialno': 1048,
                                                                                        'course': 'B.E./B.TECH/INTEGRATED COURSE THIRD YEAR SECOND SEMESTER',
                                                                                        'degree': 'B.E./B.TECH/Integrated Course',
                                                                                        'table': 'gradestructure3',
                                                                                        'appearing_year': 'APRIL 2018',
                                                                                        'Date_time': '2018-07-06 11:39:44',
                                                                                        'regno': number,
                                                                                        'revdate': '2018-07-19',
                                                                                        'revfee': 750
                                                                                        },verify=False)
        soup = BeautifulSoup(res.text,"lxml")

        details_of_student=[]
        for table in soup.find_all("table")[3:4]:
            trs = table.find_all("tr")
            details_of_student.append(trs[0].text.split(':')[1].strip())
            details_of_student.append(trs[1].text.split(':')[1].strip())

        for table in soup.find_all("table")[4:5]:
            trs = table.find_all("tr")[1:]
            for tr in trs:
                tds = tr.find_all('td')
                details_of_student.append(tds[1].text)
        return details_of_student
    except Exception as e:
        pass

#For generating marks and storing in a file.
def generate_marks(starting_no,ending_no,file_name,wb):

    ws1 = wb.create_sheet(file_name)
    ws1.append(generate_col_names(starting_no))
    for i in range(starting_no,ending_no+1):
        ws1.append(generate_individual_marks(i))

    wb.save("3-21Results.xlsx")

if __name__=='__main__':
    wb = Workbook()
    
    generate_marks(315175710001,315175710288,'CSE',wb)
    generate_marks(315175711001, 315175711184, 'IT', wb)
    generate_marks(315175714001, 315175714283, 'EEE', wb)
    generate_marks(315175712001, 315175712288, 'ECE', wb)
    generate_marks(315175720001, 315175720358, 'MECH', wb)
    generate_marks(315175708001, 315175708218, 'CIVIL', wb)
    print("done")
