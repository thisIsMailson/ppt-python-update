from pptx import Presentation
from pptx.util import Inches
# import win32com.client
import os
from pptx.util import Pt
import requests
from bs4 import BeautifulSoup
import base64
from io import BytesIO
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches
from datetime import datetime
def make_request(url, headers):
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()['data']['attributes']
    else:
        raise Exception(f"Request failed with status code {response.status_code}")
url = 'https://site-api.datocms.com/items/149001213?nested=true&version=published'
headers = {
    'X-Api-Version': '3',
    'Authorization': 'ffd35a12c12ad9223dfb0779a056ed',
    'Accept': 'application/json'
}

def getImages(data):
    upload_ids = []
    urls = []
    for item in data:
        upload_id = item['attributes']['image']['upload_id']
        response = requests.get(f'https://site-api.datocms.com/uploads/{upload_id}', headers=headers)
        if response.status_code == 200:
            urls.append(response.json()['data']['attributes']['url'])
        else:
            raise Exception(f"Request failed with status code {response.status_code}")
    return urls
    
def image_to_base64(urls):
    images = []
    for url in urls:
        response = requests.get(url)
        image_content = response.content
        base64_image = base64.b64encode(image_content).decode('utf-8')

        image_content = BytesIO(response.content)
        images.append(image_content)
    return images

def get_main_advantage(id):
    url = f'https://site-api.datocms.com/items/{id}?nested=true&version=published'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()['data']['attributes']['name']

def get_advantages(data):
    results = []
    for item in data:
        item_id = item['id']
        url = f'https://site-api.datocms.com/items/{item_id}?nested=true&version=published'
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            advantage_id = response.json()['data']['attributes']['advantage']
            advantage = response.json()['data']['attributes']['value']
            main_advantage = get_main_advantage(advantage_id)
            results.append((main_advantage, advantage))
        else:
            results.append((None, None))
    return results

try:
    product = make_request(url, headers)
    platform = product['platform']
    seg_url = f'https://site-api.datocms.com/items/{platform}?nested=true&version=published'
    segment = make_request(seg_url, headers)['name']
    cat = product['category']
    cat_url = f'https://site-api.datocms.com/items/{cat}?nested=true&version=published'
    category = make_request(cat_url, headers)['name']
    cont = product['contact_info']
    contact_url = f'https://site-api.datocms.com/items/{cont}?nested=true&version=published'
    contact = make_request(contact_url, headers)
    comp = product['company']
    company_url =  f'https://site-api.datocms.com/items/{comp}?nested=true&version=published'
    company = make_request(company_url, headers)
    div = product['division']
    division_url = f'https://site-api.datocms.com/items/{div}?nested=true&version=published'
    division = make_request(division_url, headers)
    adv_url = f'https://site-api.datocms.com/items/56982708?nested=true&version=published'
    adv = make_request(adv_url, headers)
except Exception as e:
    print(str(e))


get_advantages(product['main_advantages'])

advantage_tuple = get_advantages(product['main_advantages'])
results = advantage_tuple
# fill the rest of the results so it has at least 3 elements
if len(results) < 3:
    for i in range(len(advantage_tuple), 3):
        results.append(('', ''))
print(results)
# for advantage in product['main_advantages']:
#     print('->',advantage)

factsheet_ppt = Presentation('Motherson-Fact sheet.pptx')
slide = factsheet_ppt.slides[0]
shapes = slide.shapes
# for shape in shapes:
#     print(shape.shape_type)

slide_shapes = [
    ('title', shapes[4]), #0
    ('segment', shapes[36]), #1
    ('category', shapes[35]), #2
    ('description', shapes[5]), #3
    ('key_facts', shapes[6]), #4
    ('applications_compliancy', shapes[7]), #5
    ('intellectual_property', shapes[8]), #6
    ('first_image', shapes[1]), #7
    ('second_image', shapes[2]), #8    
    ('third_image', shapes[3]), #9
    ('img_source_1', shapes[25]), #10
    ('img_source_2', shapes[29]), #11
    ('img_source_3', shapes[27]), #12
    ('contact', shapes[21]), #13
    ('division', shapes[22]), #14
    ('email', shapes[23]), #15
    ('table', shapes[33]), #16
    ('tr_bar', shapes[38]), #17
    ('tr_4', shapes[50]), #18
]

description_shape = slide_shapes[3][1]

# print(slide_shapes[7][1].shape_type)

# for i in range(1, len(shapes)):
#     shapes[i].text = f'shape {i}'


# Change the title
title = slide_shapes[0][1].text_frame.paragraphs[0]
t_run = title.runs[0]
t_run.level = 0
t_run.font.color.rgb = RGBColor(255, 0, 0) 
title.text = BeautifulSoup(product['name'], "html.parser").get_text()

# # Change the segment
sl_segment = slide_shapes[1][1].text_frame.paragraphs[0]
run = sl_segment.runs[0]
run.font.size = Pt(8)
run.level = 0
run.text = 'Segment: '+BeautifulSoup(segment, "html.parser").get_text()

# # Change the category
sl_category = slide_shapes[2][1].text_frame.paragraphs[0]
run = sl_category.runs[0]
run.level = 0
run.text = 'Category: '+ BeautifulSoup(category, "html.parser").get_text()

# # Change the description
description = slide_shapes[3][1].text_frame.paragraphs[0]
run = description.runs[0]
run.font.size = Pt(8)
run.level = 0
run.text = BeautifulSoup(product['description'], "html.parser").get_text()

# Change the key facts
sl_keys = slide_shapes[4][1].text_frame.paragraphs[0]
run = sl_keys.runs[0]
run.font.size = Pt(7)
run.level = 0
run.text = BeautifulSoup(product['key_facts'], "html.parser").get_text()

# # Change the key applicatiin and compliancy
sl_application = slide_shapes[5][1].text_frame.paragraphs[0]
run = sl_application.runs[0]
run.font.size = Pt(7)
run.level = 0
run.text = BeautifulSoup(product['applications_compliancy'], "html.parser").get_text()

# # Change the intelectual property
sl_intelect =slide_shapes[6][1].text_frame.paragraphs[0]
run = sl_intelect.runs[0]
run.font.size = Pt(7)
run.level = 0
run.text = product['intellectual_property']

# # Change the contact
sl_contact = slide_shapes[13][1].text_frame.paragraphs[0]
run = sl_contact.runs[0]
run.font.size = Pt(6.3)
run.level = 0
c_text = 'Contact: ' + contact['name'] +'\n' +'Function: '+ contact['job_title']
run.text =  BeautifulSoup(c_text, "html.parser").get_text()

# # Change the contact
# # sl_contact_func = shapes[-3].text_frame.paragraphs[0]
# # run = sl_contact_func.runs[0]
# # run.font.size = Pt(6.3)
# # run.level = 0
# # run.text = 'Function: ' + contact['job_title']

# # shapes[-9].text = 'tr 4 text'
# # shapes[-8].text = 'Funcrtio: ' + contact['email'] division

# # Change the contact
# sl_contact = shapes[-18].text_frame.paragraphs[0]
# run = sl_contact.runs[0]
# run.font.size = Pt(6.3)
# run.level = 0
# c_text = 'Contact: ' + contact['name'] +'\n' +'Function: '+ contact['job_title']
# run.text =  BeautifulSoup(c_text, "html.parser").get_text()

# # Change the division
sl_division = slide_shapes[14][1].text_frame.paragraphs[0]
run = sl_division.runs[0]
run.font.size = Pt(6.3)
run.level = 0
d_text = 'Division: ' + division['name'] +'\n' +'Company: '+ company['name']
run.text =  BeautifulSoup(d_text, "html.parser").get_text()

# Change the Email

date_str = str(product['updated_at'])
dt = datetime.fromisoformat(date_str)

formatted_date = dt.strftime("%d-%m-%Y")

sl_email = slide_shapes[15][1].text_frame.paragraphs[0]
run = sl_email.runs[0]
run.font.size = Pt(6.3)
run.level = 0
d_text = ''+contact['email'] +'\n' +'Last modified: '+ formatted_date
run.text =  BeautifulSoup(d_text, "html.parser").get_text()

# Change the table
table = slide_shapes[16][1]

# table.style = None  # Reset the table style to default
# # table.cell(0, 0).text = 'cell 1, 1'
# # table.cell(0, 1).text = 'cell 1, 2'
# for row in table.rows:
#     for cell in row.cells:
#         cell.text_frame.text = 'New Text'
#         for paragraph in cell.text_frame.paragraphs:
#             for run in paragraph.runs:
#                 run.font.size = Pt(8)  # Set the desired font size
for i, cell in enumerate(range(15, 21)):
    curr_cell = shapes[cell].text_frame.paragraphs[0]
    run = curr_cell.runs[0]
    run.font.size = Pt(6.3)
    run.level = 0

    if i < 3:
        cell_content = results[i][0]
    else:
        cell_content = results[i - 3][1]

    run.text = BeautifulSoup(cell_content, "html.parser").get_text()


# # shapes[-20].text = 'cats 12'
images = image_to_base64(getImages(product['attachments']))

for i, image in enumerate(images):
    if i + 7 >= len(slide_shapes):
        break

    auto_shape = slide_shapes[i + 7][1]
    auto_shape.autoshape_type = MSO_SHAPE.RECTANGLE

    left = auto_shape.left
    top = auto_shape.top
    width = auto_shape.width
    height = auto_shape.height
    auto_shape_text = auto_shape.text

    slide.shapes._spTree.remove(auto_shape._element)

    new_picture = slide.shapes.add_picture(image, left, top, width, height)
    new_picture.text = auto_shape_text

ref_element = slide_shapes[10][1]._element
# print(ref_element.pre)
ref_element.addprevious(slide_shapes[7][1]._element)

# Shape to be moved
# for i in range(8, 12):
#     shape_to_remove = shapes[i]  # Replace with the shape you want to remove

#     # Access the shape's parent and remove the shape
#     parent = shape_to_remove.element.getparent()
#     parent.remove(shape_to_remove.element)

# # Get the index of the shape to be moved
# index_to_move = shapes.index(shape_to_move)

# # Remove the shape from its original position
# shapes._spTree.remove(shape_to_move._element)

# # Insert the shape at the desired position
# shapes._spTree.insert(50, shape_to_move._element)


# Update the reference to the moved shape if necessary
# if tr_bar is not None:
#     tr_bar = slide_shapes[18][1].text_frame.paragraphs[0]


# shapes[-9] tr l 



factsheet_ppt.save('fact_modified.pptx')  # Save the modified PowerPoint file

if __name__ == 'main':
    ppt()