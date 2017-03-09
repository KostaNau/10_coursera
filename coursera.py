import random
import string

import requests
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

URL_XML = 'https://www.coursera.org/sitemap~www~courses.xml'
QUANTITY = random.randrange(15, 30)
COLUMN_TITLE = [
                'title', 'Language', 'Week(s)', 'Starting date', 'Rating'
                    ]


def get_http_response(url):
    return requests.get(url).content


def parse_xlm(source_url):
    raw_xml = get_http_response(source_url)
    parser_xml = etree.XMLParser(remove_blank_text=True)
    try:
        root_xml = etree.fromstring(raw_xml, parser_xml)
        courses_urls = [
                        element.text for element in root_xml.iter()
                        if element.text is not None
                        ]
    except (etree.XMLSyntaxError) as ex:
        print('Error:', ex)
    else:
        return courses_urls


def get_random_courses(courses_urls, quantity):
    course_choices = random.sample(courses_urls, quantity)
    return course_choices


def fetch_course_data(course_url):
    title = {
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

    course_data = get_http_response(course_url)
    soup = BeautifulSoup(course_data, 'html.parser')
    course_title = soup.find(title['tag'], title['attr']).text
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
        course_weeks = None

    if rating:
        course_rating = rating.text
    else:
        course_rating = None

    return[course_title, course_language, course_weeks,
           course_date, course_rating]


def replace_none(course_data):
    for i in range(len(course_data)):
        if course_data[i] is None:
            course_data[i] = 'N/a'
    return course_data


def fill_title_column(workbook, column_title):
    sheet = workbook.active
    sheet.append(column_title)
    return workbook


def fill_data(workbook, courses_data):
    sheet = workbook.active
    for row in courses_data:
        sheet.append(row)
    return workbook


def style_workbook(workbook):

    alignment = Alignment(horizontal='center',
                          vertical='center',
                          wrap_text=True)
    sheet = workbook.active

    for column in sheet.columns:
        for cell in column:
            cell.alignment = alignment

    for cell in sheet.rows[0]:
        if cell is not None:
            cell.font = Font(bold=True)
    return workbook


def set_column_width(workbook):
    letters = string.ascii_letters
    amount_columns = 0
    width_a_cell, width_other_cell = 30, 15
    sheet = workbook.active

    for cell in sheet.rows[0]:
        if cell is not None:
            amount_columns += 1

    for letter in letters[:amount_columns]:
        if letter == 'A':
            sheet.column_dimensions[letter].width = width_a_cell
        else:
            sheet.column_dimensions[letter].width = width_other_cell
    return workbook


def save_xlx(data, column_title):
    workbook = Workbook()
    fill_title_column(workbook, column_title)
    fill_data(workbook, data)
    style_workbook(workbook)
    set_column_width(workbook)
    workbook.save('test.xlsx')


if __name__ == '__main__':
    courses_urls = parse_xlm(URL_XML)
    courses_data = []
    for course_url in get_random_courses(courses_urls, QUANTITY):
        print('Collecting course information from -', course_url)
        course_information = fetch_course_data(course_url)
        courses_data.append(replace_none(course_information))
    save_xlx(courses_data, COLUMN_TITLE)
