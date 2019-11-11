import requests
import sys
import ssl
import json
import docx
from docx import Document
from docx.shared import Inches
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

ssl._create_default_https_context = ssl._create_unverified_context
visited=set()
# write document prepare
reload(sys)
sys.setdefaultencoding('utf-8')
document = Document()
document.add_heading('Redfish Spec',0)
p = document.add_paragraph()
p.add_run('by benchin').bold = True
p.add_run()
dt={} #document table
st={} #server BMC table
def write2Document (url, jobj, next):
    document.add_heading('GET ',level = 1)
    document.add_heading('Request ',level = 2)
    document.add_paragraph(url, style = 'ListBullet')
    table = document.add_table(rows = 1,cols = 4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Name'
    hdr_cells[1].text = 'Type'
    hdr_cells[2].text = 'Read Only'
    hdr_cells[3].text = 'Description'
    for key, value in jobj.items():
        #print(key,value)
        if key == '@odata.id' and value.find("/redfish/v1") > -1:
            print("we add ",value)
            if value not in next:
                next.add(value)
            else:
                return False 
        elif key == 'Members':
            for members in value:
                if isinstance(value,list) and len(value):
                    if isinstance(value[0],dict) and value[0].has_key('@odata.id'):
                        next.add(value[0]['@odata.id'])
                        print(value)
                #break
        if isinstance(value, dict):
            write2Document (url, value, next)
            continue

        new_cells = table.add_row().cells
        new_cells[0].text = str(key)
        if type(value) is list:
            new_cells[1].text = 'Array'
        elif type(value) is int:
            new_cells[1].text = 'Integer'
        elif type(value) is dict:
            new_cells[1].text = 'Array'
        elif type(value) is bool:
            new_cells[1].text = 'Boolean'
        else:
            new_cells[1].text = 'String'
        if key.find('Action') > -1 or key.find('target') > -1 or key.find('Reset') > -1:
            new_cells[2].text = 'False'
        else:
            new_cells[2].text = 'True'

        print(key,"   ---------------    ",value)
        if key != "BiosVersion":
            new_cells[3].text = str(value)
    return True 

def go (root, url):
    if url in visited:
        return
    visited.add(url)
    uri=root+url 
    res=requests.get(str(uri), auth=('Administrator','superuser'), verify=False)
    element=json.loads(res.text)
    next = set()
    if write2Document(uri, element,next):
        st[uri] = element
    for uri in next:
        go (root, uri)

def readDoc():
    #spec = Document('/home/ben/Desktop/amidoc.docx')
    spec = docx.Document('/home/ben/Desktop/BMC_Nokia_Redfish_API_v0.4.docx')
    """for block in iter_block_items(spec):
        #print(type(block))
        if type(block) is docx.text.paragraph.Paragraph or type(block) is docx.table.Table:
            print(block.text)"""
    #iter_block_items(spec)
    for para in spec.paragraphs:
        #if para.text.find("https://{{ip}}/") > -1:
            print(para.text)
    """for section in spec.sections:
        for table in section.footer.tables:
            for i in range(len(table.rows)):
                for j in range(len(table.columns)):
                    print(table.cell(i,j).text)"""
    #print(len(spec.tables))
    found = True 
    for table in spec.tables:
        print(len(table.rows))
        print(len(table.columns))
        found = False 
        for i in range(len(table.rows)):
            for j in range(len(table.columns)):
                if table.cell(i,j).text.find ('Type URI'):
                    print(table.cell(i,j).text)
                    dt[table.cell(i,j+1).text] = table 
                    found = True 
                    break

            if found:
                break

def iter_block_items(parent):
    if isinstance(parent, docx.document.Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            print('### text is : ',child)
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            table = Table(child, parent)
            print('--------------- table man -----------------')
            for row in table.rows:
                for cell in row.cells:
                    yield iter_block_items(cell)

def main():
    readDoc()
    print("pls input your server ip ...")
    ip=sys.stdin.readline().strip('\n')
    root="https://" + str(ip)
    url="/redfish/v1/"
    print("we start from ",str(root))
    #res=requests.get(root, auth=('Administrator','superuser'))
    #res = requests.get('https://10.10.12.115/redfish/v1/', auth=('admin', 'cmb9.admin'))
    go (str(root),str(url))
    print("document table len:  ",len(dt)," server table length : ",len(st))
    for key, value in dt.items():
        print(key, " : ", value)
    print('=======================================================')
    for key, value in st.items():
        print(key, " : ", value)
    #res=requests.get('https://10.10.12.115/redfish/v1/', auth=('Administrator','superuser'), verify=False)
   # print(res.text)
    document.save('/home/ben/Desktop/testspec.docx')
if __name__ == '__main__':
    main()
