from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
from selenium import webdriver
import codecs

import json


target_url = 'https://www.newyorker.com/crossword/puzzles-dept/2019/04/01'

def simple_get(url):
    """
    Attempts to get the content at `url` by making an HTTP GET request.
    If the content-type of response is some kind of HTML/XML, return the
    text content, otherwise return None.
    """
    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during requests to {0} : {1}'.format(url, str(e)))
        return None


def is_good_response(resp):
    """
    Returns True if the response seems to be HTML, False otherwise.
    """
    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200 
            and content_type is not None 
            and content_type.find('html') > -1)


def log_error(e):
    """
    It is always a good idea to log errors. 
    This function just prints them, but you can
    make it do anything.
    """
    print(e)




ny_raw_html = simple_get(target_url)
ny_html = BeautifulSoup(ny_raw_html, 'html.parser')



crossword_iframe = ny_html.find("iframe", {"id": "crossword"})
cdn_source = crossword_iframe['data-src']


driver = webdriver.Chrome()
driver.get(cdn_source)

crossword = driver.find_element_by_class_name('crossword').get_attribute('innerHTML')
aclues = driver.find_element_by_class_name('aclues').get_attribute('innerHTML')
dclues = driver.find_element_by_class_name('dclues').get_attribute('innerHTML')

driver.close()


crossword_html = BeautifulSoup(crossword, 'html.parser')
puzzle_array = [] 
current_row = []
for div in crossword_html.findAll("div"):
    # print(div)
    if 'endRow' in div['class']:
        puzzle_array.append(current_row)
        current_row = []
    elif not div.find("img", {"src":"images/black1px.png"}) is None:
        current_row.append("#")
    elif len(div.text.strip()) > 0:
        current_row.append(int(div.text))
    else:
        current_row.append(0)



aclues_html = BeautifulSoup(aclues, 'html.parser')
aclues_list = [] 
for div in aclues_html.findAll("div"):
    if div.has_attr('class') and 'clueDiv' in div['class']:
        cluenum = int(div.find("div", {"class":"clueNum"}).text)
        cluetext = div.find("div", {"class":"clue"}).text
        aclues_list.append([cluenum, cluetext])

dclues_html = BeautifulSoup(dclues, 'html.parser')
dclues_list = [] 
for div in dclues_html.findAll("div"):
    if div.has_attr('class') and 'clueDiv' in div['class']:
        cluenum = int(div.find("div", {"class":"clueNum"}).text)
        cluetext = div.find("div", {"class":"clue"}).text
        dclues_list.append([cluenum, cluetext])


puzzle = {}

puzzle["version"] = target_url
puzzle["kind"] = [target_url]

puzzle["puzzle"] = puzzle_array

puzzle["clues"] = {"Across":aclues_list, "Down":dclues_list}


puzzle_serialized = json.dumps(puzzle)

f_out = codecs.open("out.ipuz", "w+")
f_out.write(puzzle_serialized)
f_out.close()

def to_sup(s):
    sups = {u'0': u'\u2070',
            u'1': u'\xb9',
            u'2': u'\xb2',
            u'3': u'\xb3',
            u'4': u'\u2074',
            u'5': u'\u2075',
            u'6': u'\u2076',
            u'7': u'\u2077',
            u'8': u'\u2078',
            u'9': u'\u2079'}

    return ''.join(sups.get(char, char) for char in s)  # lose the list comprehension


import xlwt 
from xlwt import Workbook 
  
# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet = wb.add_sheet("crossword") 
  

x = len(puzzle_array)
y = len(puzzle_array[0])

tall_style = xlwt.easyxf('font:height 500;') # 36pt

for i in range(x):
    sheet.row(i).set_style(tall_style) 
for j in range(y):
    sheet.col(j).width = 500

for i in range(x):
    for j in range(y):
        target = str(puzzle_array[i][j])
        if "#" in target:
            target = ''
            style = xlwt.easyxf('pattern: pattern solid, fore_colour black;')
            sheet.write(i, j, target , style) 
        elif "0" == target:
            pass
        else:
            target = to_sup(target.strip())
            style = xlwt.easyxf('font: height 100;')
            sheet.write(i, j, target , style) 


row_num  = 0
style = xlwt.easyxf('font: height 100;')

sheet.write(row_num, x + 1, "Clues across:" , style) 
row_num += 1

for z in aclues_list:
    sheet.write(row_num, x + 1, str(z[0]) , style) 
    sheet.write(row_num, x + 2, str(z[1]) , style) 
    row_num += 1

sheet.write(row_num, x + 1, "Clues down:" , style) 
row_num += 1

for z in dclues_list:
    sheet.write(row_num, x + 1, str(z[0]) , style) 
    sheet.write(row_num, x + 2, str(z[1]) , style) 
    row_num += 1


wb.save('out.xls') 
