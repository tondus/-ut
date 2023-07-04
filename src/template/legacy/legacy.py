from src.lib.template import *

def add_data_to_table(table , data ,config,ricefw,pic,text=None):
    set_checkbox_from_data(data,table)
    set_other_from_data(data,table)
    set_source_target_from_data(data,table)
 
    row_content = table.row_cells(6)
    paragraph_pic = row_content[0].add_paragraph()
    # paragraph_process = row_content[0].add_paragraph()
    paragraph_body_line1 = row_content[0].add_paragraph()
    
    set_picture_from_data(paragraph_pic,pic)
    # set_process_from_data(paragraph_process,data)  
    set_text_from_data_oneline(paragraph_body_line1,text)

    return table

def get_pic(data,config):
    pic = None 
    text = None
    
    if data[col_asynchronus]:
        if data[col_inbound] : 
            pic = root_images + "legacy/async_in.png"
            text = 'Legacy ทำการส่งข้อมูลผ่าน Web Services ให้ API Management เพื่อส่งต่อ SAP PI ทำการ Mapping Field เพื่อเรียกใช้งาน BAPI Function ในระบบ SAP S/4HANA ผ่าน ABAP Proxy'
        else:
            pic = root_images + "legacy/async_out.png"
            text = 'ระบบ SAP S/4HANA ทำการส่งข้อมูลผ่าน ABAP Proxy ส่งต่อ SAP PI ทำการ Mapping Field เพื่อส่งข้อมูลไปยัง API Management เพื่อส่งกลับไปยัง Legacy ผ่าน Web Service'
    else : 
        if data[col_inbound] : 
            pic = root_images + "legacy/sync_in.png"
            text = 'Legacy ทำการส่งข้อมูลผ่าน Web Services ให้ API Management เพื่อส่งต่อ SAP PI ทำการ Mapping Field เพื่อเรียกใช้งาน BAPI Function ในระบบ SAP S/4HANA ผ่าน ABAP Proxy  หลังจากที่ SAP S/4HANA ทำการ Process ข้อมูลเรียบร้อยแล้ว จะส่งผลการทำงานผ่าน SAP PI ส่งต่อ API Management กลับทางเส้นทางเดิมเพื่อส่งข้อมูลกลับไปยัง Legacy'
        else:
            pic = root_images + "legacy/sync_out.png"
            text = 'SAP S/4HANA ทำการส่งข้อมูลผ่าน ABAP Proxy มายัง SAP PI เพื่อ Mapping Field ไปให้ API Management ส่งต่อไปยัง Legacy โดยผ่าน Web Service หลังจากที่ Legacy ทำการ Process ข้อมูลเรียบร้อยแล้ว จะส่งผลการทำงานผ่าน Web Service มายัง API Management ส่งต่อให้ SAP PI กลับทางเส้นทางเดิมเพื่อส่งข้อมูลกลับไปยัง SAP S/4HANA'
    return pic , text
