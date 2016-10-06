import xml.etree.ElementTree as ET
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import json


COURSE_COUNT = 5
CELLS_NAME = ('Name', 'Language', 'Start date', 'Duration', 'Rating')


def get_courses_list():
    node_prefix = '{http://www.sitemaps.org/schemas/sitemap/0.9}'
    xml_site = requests.get('https://www.coursera.org/sitemap~www~courses.xml').content
    xml_tree_root = ET.fromstring(xml_site.decode('utf-8'))
    xml_courses = xml_tree_root.findall('{prefix}url'.format(prefix=node_prefix))
    courses_list = [course.find('{prefix}loc'.format(prefix=node_prefix)).text for course in xml_courses]
    return courses_list
    

def get_course_info(course_link):
    course_html = requests.get(course_link)
    if course_html.url not in course_link:
        return
    soup = BeautifulSoup(course_html.text, 'html.parser')
    basic_info = soup.find('table', 'basic-info-table')
    course_name = soup.find('div', 'title').text
    course_lang = ' '.join(basic_info.find(string='Language').next_element.stripped_strings)
    course_duration = basic_info.find(string='Commitment')
    if course_duration:
        course_duration = ' '.join(course_duration.next_element.stripped_strings)
    course_start_date = soup.select('.rc-CourseGoogleSchemaMarkup')
    if course_start_date:
        course_start_date = json.loads(course_start_date[0].text)['hasCourseInstance'][0]['startDate']
    course_rating = soup.select('.ratings-text.bt3-visible-xs')
    if course_rating:
        course_rating = course_rating[0].text.split(' ')[0]
    full_info = [course_name, course_lang, course_start_date, course_duration, course_rating]
    print('Getting %s' % course_name)
    return full_info


def output_courses_info_to_xlsx(filepath, courses):
    wb = Workbook()
    ws = wb.active
    ws.title = '20 random courses from Coursera'
    for name, cell in enumerate(list(ws['A1:E1'])[0]):
        cell.value = CELLS_NAME[name]
    ###THIS PART NOT WORKING
    for course in ws.iter_rows(min_row=2, max_col=5, max_row=COURSE_COUNT+1):
        for cell_name, cell in enumerate(course):
            if courses[cell_name]:
                cell.value = courses[cell_name]
    ###
    wb.save(filepath)


if __name__ == '__main__':
    courses_list = get_courses_list()
    courses_info = []
    while len(courses_info) != COURSE_COUNT:
        courses_info.append(get_course_info(courses_list.pop()))
    output_courses_info_to_xlsx('Coursera.xlsx', courses_info)
