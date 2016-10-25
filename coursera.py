import json
import random
from datetime import datetime

import requests
import babel
from lxml import etree as ET
from lxml import html as HT
from openpyxl import Workbook
from openpyxl.styles import Font


COURSE_COUNT = 100

COURSERA_FEED = 'https://www.coursera.org/sitemap~www~courses.xml'
XML_PREFIX = '{http://www.sitemaps.org/schemas/sitemap/0.9}'

COLUMN_NAMES = ('Name',
    'Language',
    'Launch date',
    'Upcoming Session',
    'Start date',
    'Duration',
    'Rating')


def get_courses_urls():
    response = requests.get(COURSERA_FEED)
    xml_root = ET.fromstring(response.content)
    xml_url_nodes = xml_root.iterfind('.//%sloc' % XML_PREFIX)
    courses_urls = [clean_url.text for clean_url in xml_url_nodes]
    return courses_urls


def get_course_info(course_url):
    course_slug = course_url.split('/').pop()
    payload = {
        'q': 'slug',
        'slug': course_slug,
        'fields': 'workload,\
            primaryLanguages,\
            startDate,\
            plannedLaunchDate,\
            upcomingSessionStartDate,\
            courseDerivatives.v1(averageFiveStarRating)',
        'includes': 'courseDerivatives'
    }
    api_response = requests.get(
        'https://api.coursera.org/api/courses.v1',
        params=payload,
    )
    api_json = json.loads(api_response.text)
    api_elements = api_json['elements'].pop()
    api_rating = api_json['linked']['courseDerivatives.v1'].pop()
    
    name = api_elements['name']
    workload = get_workload(api_elements)
    language = get_languages(api_elements)
    launch_date = get_launch_date(api_elements)
    rating = get_rating(api_rating)
    
    course_info = [
        name,
        language,
        launch_date['launch_date'],
        launch_date['upcoming_date'],
        launch_date['start_date'],
        workload,
        rating
    ]
    return course_info


def get_workload(json_data):
    workload = json_data['workload']
    if workload:
        return workload 
    else:
        return 'Unknown'


def get_languages(json_data):
    lang_code = json_data['primaryLanguages'].pop()
    language = babel.Locale.parse(lang_code, sep='-').english_name
    return language


def get_launch_date(json_data):
    try:
        launch_normal = json_data['plannedLaunchDate']
    except KeyError:
        launch_normal = None
    try:
        upcoming = json_data['upcomingSessionStartDate']/1000
        upcoming_normal = datetime.fromtimestamp(upcoming).strftime('%B %d %Y')
    except KeyError:
        upcoming_normal = None
    
    try:
        start_date = json_data['startDate']/1000
        start_normal = datetime.fromtimestamp(start_date).strftime('%B %d %Y')
    except KeyError:
        start_normal = None
    
    return {
        'launch_date': launch_normal,
        'upcoming_date': upcoming_normal,
        'start_date': start_normal,
    }


def get_rating(json_data):
    try:
        return round(json_data['averageFiveStarRating'],1)
    except KeyError:
        return 'Not rated'


def output_courses_info_to_xlsx(filepath, courses):
    wb = Workbook()
    ws = wb.active
    ws.title = '20 random courses from Coursera'
    for name, cell in enumerate(list(ws['A1:G1'])[0]):
        cell.value = COLUMN_NAMES[name]
        cell.font = Font(bold=True)
    ws.column_dimensions['A'].width = 30
    for row in ws.iter_rows(min_row=2, max_col=7, max_row=COURSE_COUNT+1):
        for cell_name, cell in enumerate(row):
            if courses[cell.row-2][cell_name]:
                cell.value = courses[cell.row-2][cell_name]
    wb.save(filepath)


if __name__ == '__main__':
    #DEBUG
    #course_info = get_course_info('https://www.coursera.org/learn/zhichang-suyang')
    #print(course_info)
    ###################
    print('Getting course list...')
    courses_list = get_courses_urls()
    print('Done!')
    courses_info = []
    print('Getting %d courses info...' % COURSE_COUNT)
    while len(courses_info) != COURSE_COUNT:
        random_index = courses_list.index(random.choice(courses_list))
        random_course = courses_list.pop(random_index)
        #random_course = courses_list.pop(0)
        print('Get: %s' % random_course)
        course_info = get_course_info(random_course)
        if course_info:
            courses_info.append(course_info)
    print('Done!')
    print('Making excel file...')
    output_courses_info_to_xlsx('Coursera.xlsx', courses_info)
    print('Finish!')


'''
def get_course_info_by_lxml(course_url):
    response = requests.get(course_url)
    if response.status_code != 200:
        return
    
    #DEBUG
    #with open('scratch.html','w') as scratch:
        #scratch.write(response.text)

    html_root = HT.fromstring(response.content)

    title_tag = html_root.find_class('title').pop()
    title = title_tag.text

    try:
        startDate_block = html_root.find_class('rc-CourseGoogleSchemaMarkup').pop()
        startDate_json = json.loads(startDate_block.text_content()) 
        startDate = startDate_json['hasCourseInstance'].pop()['startDate'] 
    except IndexError:
        startDate = 'Already started'
    except KeyError:
        startDate = 'No upcomming session'

    basic_info = html_root.find_class('basic-info-table').pop()
    lang_tag = basic_info.find_class('cif-language').pop().getparent().getnext()
    language = lang_tag.text_content()
    try:
        workload_tag = basic_info.find_class('cif-clock').pop().getparent().getnext()
        workload = workload_tag.text
    except IndexError:
        workload = 'No info'
    try:
        rating_block = basic_info.find_class('ratings-text bt3-visible-xs').pop()
        rating = rating_block.text
    except IndexError:
        rating = 'Not rated'

    course_info = [title, language, startDate, workload, rating]
    return course_info
'''





'''
    try:
        launch_normal = datetime.strptime(launch_date, '%B %Y').strftime('%B %Y')
    except ValueError:
        fix_day = ''.join(launch_date.split('th')).strip()
        launch_normal = datetime.strptime(fix_day, '%B %d, %Y').strftime('%B %d %Y')

'''