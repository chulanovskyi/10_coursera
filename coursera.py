import xml.etree.ElementTree as ET
import requests
from bs4 import BeautifulSoup
import openpyxl


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
        print('Course not available')
        return
    soup = BeautifulSoup(course_html.text, 'html.parser')
    basic_info = soup.find('table', 'basic-info-table')
    course_name = soup.find('h1', 'course-name-text')
    course_lang = ' '.join(basic_info.find(string='Language').next_element.stripped_strings)
    course_duration = ' '.join(basic_info.find(string='Commitment').next_element.stripped_strings)
    print('-'*20)
    print('-'*20)


def output_courses_info_to_xlsx(filepath):
    pass


if __name__ == '__main__':
    course_list = get_courses_list()
    for course in range(2):
        get_course_info(course_list[course])
        print('-'*30)
        print(course_list[course])
