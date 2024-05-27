import requests
from bs4 import BeautifulSoup
from collections import defaultdict
import xlsxwriter



def fetch_exam_info(url, exam_info, topic):

    # Send en GET-forespørsel til URL-en med det angitte språket
    response = requests.get(url)
    
    # Se om forespørselen var vellykket
    if response.status_code != 200:
        print("Feil ved henting av nettsiden.")
        return
    
    # Parse HTML-innholdet
    soup = BeautifulSoup(response.content, 'html.parser')
    
       
    # Finn eksamensdato og tid
    exam_date_section = soup.find('p', class_='exam-date')
    if exam_date_section:
        
        exam_time_text = exam_date_section.text.strip("")
        exam_time_text = exam_time_text.strip("\n").split(",")
        topic_info= "".join(exam_time_text).split()
        
        date = "".join(topic_info[1:3])
        start_time = "".join(topic_info[4])
        duration = " ".join(topic_info[5:]).strip(".")
        duration = duration.lstrip("(")
        
        exam_info[topic] = {"Date":date, "StartTime":start_time,"Duration":duration.rstrip(")")}
    
    return exam_info


def fetch_course_codes(url):
    # Hent HTML-innholdet fra nettsiden
    response = requests.get(url)
    if response.status_code != 200:
        print("Feil ved henting av nettsiden. Ingen semesterside lagt til")
        return []

    # Parse HTML-innholdet med BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Finn alle emnekoder
    course_codes = []
    for row in soup.select('table tr'):
        columns = row.find_all('td')
        if columns:
            course_code = columns[0].text.strip()
            course_codes.append(course_code.split()[0])

    return course_codes



def main():
    
    bachelor = "https://www.uio.no/studier/emner/matnat/ifi/?filter.level=bachelor&filter.semester=v24"
    master = "https://www.uio.no/studier/emner/matnat/ifi/?filter.level=master&filter.semester=h24"
    
    courses = fetch_course_codes(bachelor)
    topic_overview = defaultdict(str)
    
    data = [
        ['Emnekode', 'Eksamendato', "Tidspunkt", "Varighet"]
    ]
    
    for emne in courses:
        url = f"https://www.uio.no/studier/emner/matnat/ifi/{emne}/v24/eksamen/index.html"
        fetch_exam_info(url, topic_overview, emne)

  
    for key in topic_overview:
        topic_info = [key, topic_overview[key]["Date"], topic_overview[key]["StartTime"], topic_overview[key]["Duration"]]
        
        data.append(topic_info)
    
    workbook = xlsxwriter.Workbook('exam_calendar.xlsx')
    worksheet = workbook.add_worksheet()
    
    for row_idx, row_data in enumerate(data):
        for col_idx, cell_data in enumerate(row_data):
            worksheet.write(row_idx, col_idx, cell_data)

    # Lukk arbeidsboken
    workbook.close()
if __name__ == "__main__":
    main()

