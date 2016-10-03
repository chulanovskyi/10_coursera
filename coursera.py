import xml.etree.ElementTree as ET
import requests
import bs4
import openpyxl


def get_courses_list():
    node_prefix = '{http://www.sitemaps.org/schemas/sitemap/0.9}'
    xml_site = requests.get('https://www.coursera.org/sitemap~www~courses.xml').content
    xml_tree_root = ET.fromstring(xml_site.decode('utf-8'))
    xml_courses = xml_tree_root.findall('{prefix}url'.format(prefix=node_prefix))
    courses_link_list = [course.find('{prefix}loc'.format(prefix=node_prefix)).text for course in xml_courses]
    return courses_link_list
    

def get_course_info(course_link):
    course_html = requests.get(course_link)
    print(course_html)


def output_courses_info_to_xlsx(filepath):
    pass


if __name__ == '__main__':
    course_list = get_courses_list()
    get_course_info(course_list[0])
