from src.lib.template import *

def add_data_to_table(table , data ,config,ricefw,pic,text=None):
    set_checkbox_from_data(data,table)
    set_other_from_data(data,table)
    set_source_target_from_data(data,table)
    
    row_content = table.row_cells(6)
    paragraph_pic = row_content[0].add_paragraph()
    # paragraph_process = row_content[0].add_paragraph()
 
    
    set_picture_from_data(paragraph_pic,pic)
    # set_process_from_data(paragraph_process,data)  
    print('tas',data[col_inbound])
    for t in text:
        paragraph_body_line1 = row_content[0].add_paragraph()
        set_text_from_data_bullet(paragraph_body_line1,t)
   
        
    return table

def get_pic(data):
    pic = None
    text = None
    if data[col_inbound]:
        pic = root_images + r'\tas\tas_in.png' 
        text = ['•	TAS ทำการวางไฟล์ไว้ที่ Server แยกวางใน Folder ตามแต่ละ Plant',
                '•	ESB ทำการย้ายไฟล์ไปวางที่ SFTP Server แยกวางใน Folder ตามแต่ละ Plant',
                '•	SAP PI ทำการหยิบไฟล์ที่ SFTP Server ที่แยกวางใน Folder ตามแต่ละ Plant เพื่อทำ Mapping field call IDoc เข้า SAP S/4HANA']
    else:
        pic = root_images + r'\tas\tas_out.png' 
        text = ['•	SAP S/4HANA ทำการส่ง IDoc มายัง SAP PI',
                '•	SAP PI ทำการ Mapping field IDoc ให้เป็น text file และวางไปที่ SFTP Server แยกวางใน Folder ตามแต่ละ Plant',
                '•	ESB ทำการย้ายไฟล์จาก SFTP Server ไปยัง Server TAS โดยแยกวางใน Folder ตามแต่ละ Plant',
                '•	TAS หยิบไฟล์จาก Server เพื่อทำการ Process ข้อมูล']
    return pic,text