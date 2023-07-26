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
    if 'HR' in data[col_source_system].upper() or 'HR' in data[col_target_system].upper():
        pic = root_images + "rfc/hr/sync_in.png"
        print('HR',data[col_ricef])
        text = 'Legacy ทำการส่งข้อมูลผ่าน Web Services ให้ API Management เพื่อส่งต่อ SAP PI ทำการ Mapping Field เพื่อเรียกใช้งาน RFC ในระบบ SAP S/4HANA (HR) หลังจากที่ SAP S/4HANA (HR) ทำการ Process ข้อมูลเรียบร้อยแล้ว จะส่งผลการทำงานผ่าน SAP PI ส่งต่อ API Management กลับทางเส้นทางเดิมเพื่อส่งข้อมูลกลับไปยัง Legacy'
    elif 'BW' in data[col_source_system].upper() or 'BW' in data[col_target_system].upper():
        pic = root_images + "rfc/bw/sync_in.png"
        print('BW',data[col_ricef])
        text = 'Legacy ทำการส่งข้อมูลผ่าน Web Services ให้ API Management เพื่อส่งต่อ SAP PI ทำการ Mapping Field เพื่อเรียกใช้งาน RFC ในระบบ BW/4HANA หลังจากที่ BW/4HANA ทำการ Process ข้อมูลเรียบร้อยแล้ว จะส่งผลการทำงานผ่าน SAP PI ส่งต่อ API Management กลับทางเส้นทางเดิมเพื่อส่งข้อมูลกลับไปยัง Legacy'
    else:
        pic = root_images + "rfc/sap/sync_in.png"
        print('SAP',data[col_ricef])
        text = 'Legacy ทำการส่งข้อมูลผ่าน Web Services ให้ API Management เพื่อส่งต่อ SAP PI ทำการ Mapping Field เพื่อเรียกใช้งาน RFC ในระบบ SAP S/4HANA หลังจากที่ SAP S/4HANA ทำการ Process ข้อมูลเรียบร้อยแล้ว จะส่งผลการทำงานผ่าน SAP PI ส่งต่อ API Management กลับทางเส้นทางเดิมเพื่อส่งข้อมูลกลับไปยัง Legacy'
    return pic , text
