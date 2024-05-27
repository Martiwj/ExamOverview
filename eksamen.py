import requests
from bs4 import BeautifulSoup
from collections import defaultdict
from datetime import datetime
import xlsxwriter



def fetch_exam_info(url, exam_info, emne):
    # Send en GET-forespørsel til URL-en
    språkkode = 'no'  # for norsk, 'en' for engelsk, osv.

    # Definer forespørselshodet med språkinnstillingen
    headers = {'Accept-Language': språkkode}

    # Send en GET-forespørsel til URL-en med det angitte språket
    response = requests.get(url, headers=headers)
    
    # Sjekk om forespørselen var vellykket
    if response.status_code != 200:
        print("Feil ved henting av nettsiden.")
        return
    
    # Parse HTML-innholdet
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Finn informasjon om eksamen
   
    
    # Finn eksamensdato og tid
    exam_date_section = soup.find('p', class_='exam-date')
    if exam_date_section:
        
        exam_time_text = exam_date_section.text.strip("")
        exam_time_text = exam_time_text.strip("\n").split(",")
        emne_in= "".join(exam_time_text).split()
        
        date = "".join(emne_in[1:3])
        start_time = "".join(emne_in[4])
        duration = " ".join(emne_in[5:]).strip(".")
        duration = duration.lstrip("(")
        
        exam_info[emne] = {"Date":date, "StartTime":start_time,"Duration":duration.rstrip(")")}
        
    
    # Finn eksamenssted
    """  exam_place_section = soup.find('p', class_='exam-place-info')
        if exam_place_section:
            places = [place.strip() for place in exam_place_section.stripped_strings]
            exam_info['Sted'] = places
        """
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
    emne_oversikt = defaultdict(str)
    
    data = [
        ['Emnekode', 'Eksamendato', "Tidspunkt", "Varighet"]
    ]
    
    for emne in courses:
        url = f"https://www.uio.no/studier/emner/matnat/ifi/{emne}/v24/eksamen/index.html"
        fetch_exam_info(url, emne_oversikt, emne)

  
    for key in emne_oversikt:
        emne_info = [key, emne_oversikt[key]["Date"], emne_oversikt[key]["StartTime"], emne_oversikt[key]["Duration"]]
        
        data.append(emne_info)
    
    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet()
    
    

    for row_idx, row_data in enumerate(data):
        for col_idx, cell_data in enumerate(row_data):
            worksheet.write(row_idx, col_idx, cell_data)

    # Lukk arbeidsboken
    workbook.close()
if __name__ == "__main__":
    main()

