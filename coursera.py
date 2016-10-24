import json
import random

import requests
from lxml import etree as ET
from lxml import html as HT
from openpyxl import Workbook
from openpyxl.styles import Font


COURSE_COUNT = 20
COURSERA_FEED = 'https://www.coursera.org/sitemap~www~courses.xml'
XML_PREFIX = '{http://www.sitemaps.org/schemas/sitemap/0.9}'

CELLS_NAME = ('Name', 'Language', 'Start date', 'Duration', 'Rating')
COL_WIDTH_L = 50
COL_WIDTH_S = 25


def get_courses_urls():
    response = requests.get(COURSERA_FEED)
    xml_root = ET.fromstring(response.content)
    xml_url_nodes = xml_root.iterfind('.//%sloc' % XML_PREFIX)
    courses_urls = [clean_url.text for clean_url in xml_url_nodes]
    return courses_urls


def get_course_info(course_url):
    response = requests.get(course_url)
    html_root = HT.fromstring(response.content)

    title_tag = html_root.find_class('title').pop()
    title = title_tag.text

    basic_info = html_root.find_class('basic-info-table').pop()
    workload_tag = basic_info.find_class('cif-clock').pop().getparent().getnext()
    workload = workload_tag.text

    lang_tag = basic_info.find_class('cif-language').pop().getparent().getnext()
    language = lang_tag.text_content()

    startDate_block = html_root.find_class('rc-CourseGoogleSchemaMarkup').pop()
    startDate_json = json.loads(startDate_block.text_content()) 
    startDate = startDate_json['hasCourseInstance'].pop()['startDate'] 

    try:
        rating_block = basic_info.find_class('ratings-text bt3-visible-xs').pop()
        rating = rating_block.text
    except IndexError:
        rating = 'Not rated'

    course_info = [title, language, startDate, workload, rating]
    return course_info    


def output_courses_info_to_xlsx(filepath, courses):
    wb = Workbook()
    ws = wb.active
    ws.title = '20 random courses from Coursera'
    for name, cell in enumerate(list(ws['A1:E1'])[0]):
        cell.value = CELLS_NAME[name]
        cell.font = Font(bold=True)
    ws.column_dimensions['A'].width = COL_WIDTH_L
    ws.column_dimensions['B'].width = COL_WIDTH_L
    ws.column_dimensions['C'].width = COL_WIDTH_S
    ws.column_dimensions['D'].width = COL_WIDTH_S
    ws.column_dimensions['E'].width = COL_WIDTH_S
    for row in ws.iter_rows(min_row=2, max_col=5, max_row=COURSE_COUNT+1):
        for cell_name, cell in enumerate(row):
            if courses[cell.row-2][cell_name]:
                cell.value = courses[cell.row-2][cell_name]
    wb.save(filepath)


if __name__ == '__main__':
    print('Getting course list...')
    courses_list = get_courses_urls()
    print('Done!')
    courses_info = []
    print('Getting %d courses info...' % COURSE_COUNT)
    while len(courses_info) != COURSE_COUNT:
        random_index = courses_list.index(random.choice(courses_list))
        random_course = courses_list.pop(random_index)
        course_info = get_course_info(random_course)
        if course_info:
            courses_info.append(course_info)
    print('Done!')
    print(len(courses_info))
    #print('Making excel file...')
    #output_courses_info_to_xlsx('Coursera.xlsx', courses_info)
    #print('Finish!')
