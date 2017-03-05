import random

import requests
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font


LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
COURSERA_XML = 'https://www.coursera.org/sitemap~www~courses.xml'
QUANTITY = random.randrange(15, 30)
COLUMN_TITLE = [
                'Name', 'Language', 'Week(s)', 'Starting date', 'Rating'
                    ]


def parse_xml_content(xml_url: str) -> list:
    raw_xml = requests.get(xml_url).content
    parser_xml = etree.XMLParser(remove_blank_text=True)
    try:
        root_xml = etree.fromstring(raw_xml, parser_xml)
        elements_content = [
                        element.text for element in root_xml.iter()
                        if element.text is not None
                        ]
    except (etree.XMLSyntaxError) as ex:
        print('Error:', ex)
    else:
        return elements_content


def get_random_courses(courses_stack: list, quantity: int) -> tuple:
    course_choices = random.sample(courses_stack, quantity)
    return course_choices


def fetch_course_info(course_link):
    name = {
            'tag': 'h1',
            'attr': {'class': 'title display-3-text'}
            }
    language = {
            'tag': 'div',
            'attr': {'class': 'language-info'}
            }
    date = {
            'tag': 'div',
            'attr': {'class': 'startdate rc-StartDateString caption-text'}
            }
    week = {
            'tag': 'div',
            'attr': {'class': 'week'},
            'alt_tag': 'td',
            'alt_attr': {'class': 'td-data'}
            }
    rating = {
            'tag': 'div',
            'attr': {'class': 'ratings-text bt3-visible-xs'}
            }

    course_data = requests.get(course_link).content
    soup = BeautifulSoup(course_data, 'html.parser')
    course_name = soup.find(name['tag'], name['attr']).text
    course_language = soup.find(language['tag'], language['attr']).text
    course_date = soup.find(date['tag'], date['attr']).text

    weeks = soup.find(week['tag'], week['attr'])
    alt_weeks = soup.find(week['alt_tag'], week['alt_attr'])
    rating = soup.find(rating['tag'], rating['attr'])

    if weeks:
        course_length = len(weeks)
        course_weeks = '{} week(s) of study'.format(course_length)
    elif 'Week' in alt_weeks.text:
        course_weeks = alt_weeks.text
    else:
        course_weeks = 'N/a'
        """
        Attention for the mentor: In my opinion str is much better than None
        for view in Excel. All the same we do it for GUI, instead console
        output and perhaps for not experienced user.
        """

    if rating:
        course_rating = rating.text
    else:
        course_rating = 'N/a'

    return(course_name, course_language, course_weeks,
           course_date, course_rating)


def create_xlsx(courses_data: tuple, column_title: list):

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(column_title)
    for row in courses_data:
        sheet.append(row)
    return workbook


def style_pyxl_object(raw_workbook):

    length_column_a = 30
    length_column_other = 15

    al = Alignment(horizontal='center',
                   vertical='center',
                   wrap_text=True)

    workbook = raw_workbook
    sheet = workbook.active

    # alignment for every cell
    for column in sheet.columns:
        for cell in column:
            cell.alignment = al

    """
    Add bold font for every column title cell and calculate amount of used
    column (needs for stetting custom column width, see bellow)
    """
    amount_columns = 0
    for cell in sheet.rows[0]:
        if cell is not None:
            cell.font = Font(bold=True)
            amount_columns += 1

    # Add custom column width
    for letter in LETTERS[:amount_columns]:
        if letter == 'A':
            sheet.column_dimensions[letter].width = length_column_a
        else:
            sheet.column_dimensions[letter].width = length_column_other
    return workbook


if __name__ == '__main__':
    courses = parse_xml_content(COURSERA_XML)
    courses_info = []
    for course in get_random_courses(courses, QUANTITY):
        print('Collecting course information from -', course)
        courses_info.append(fetch_course_info(course))
    """
    Of course list comprhenesion is better and faster, with it I lost output
    in console statement information about current process.
    """
    # courses_info = [
    #         fetch_course_info(course) for course in
    #         get_random_courses(courses, QUANTITY)
    #                     ]
    raw_pyxl = create_xlsx(courses_info, COLUMN_TITLE)
    pyxl_object = style_pyxl_object(raw_pyxl)
    pyxl_object.save('test.xlsx')
