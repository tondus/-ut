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
    if data[col_proxy]:
        paragraph_body_line1 = row_content[0].add_paragraph()
        set_text_from_data_oneline(paragraph_body_line1,text[0])
    else:
        for t in text:
            paragraph_body_line1 = row_content[0].add_paragraph()
            set_text_from_data_bullet(paragraph_body_line1,t)

    return table

def get_pic(data,config):
    pic = None 
    if data[col_other1] == 'File':
        if data[col_inbound]:
            pic = root_images + "bw/file_in.png"
            text = ['• ESB ทำการย้ายไฟล์จาก File Server ไปวางที่ SFTP Server ',
                    '• SAP PI ทำการหยิบไฟล์ที่ SFTP Server เพื่อวางไว้ที่ AL11 บน SAP BW/4HANA',
                    '• คง AL11 Path, File server และ Path บน File Server ต้นทางไว้ตาม AS-IS'] 
        else: 
            pic = root_images + "bw/file_out.png"
            text = ['•	SAP BW/4HANA ทำการวางไฟล์ไว้ที่ AL11',
                    '•	SAP PI ทำการหยิบไฟล์ที่ AL11 บน SAP BW/4HANA ไปวางที่ SFTP Server',
                    '•	ESB ทำการย้ายไฟล์จาก SFTP Server ไปยัง File Server',
                    '•	คง AL11 Path, File server และ Path บน File Server ปลายทางไว้ตาม AS-IS'] 
    elif data[col_proxy]:
        if data[col_inbound]:
            pic = root_images + "bw/sync_in.png"
            text = ['Legacy ทำการส่งข้อมูลผ่าน Web Services ให้ API Management เพื่อส่งต่อ SAP PI ทำการ Mapping Field เพื่อเรียกใช้งาน BAPI Function ในระบบ BW/4HANA ผ่าน ABAP Proxy  หลังจากที่ BW/4HANA ทำการ Process ข้อมูลเรียบร้อยแล้ว จะส่งผลการทำงานผ่าน SAP PI ส่งต่อ API Management กลับทางเส้นทางเดิมเพื่อส่งข้อมูลกลับไปยัง Legacy'] 
        else: 
            pic = root_images + "bw/sync_out.png"
            text = ['Legacy ทำการส่งข้อมูลผ่าน Web Services ให้ API Management เพื่อส่งต่อ SAP PI ทำการ Mapping Field เพื่อเรียกใช้งาน BAPI Function ในระบบ BW/4HANA ผ่าน ABAP Proxy  หลังจากที่ BW/4HANA ทำการ Process ข้อมูลเรียบร้อยแล้ว จะส่งผลการทำงานผ่าน SAP PI ส่งต่อ API Management กลับทางเส้นทางเดิมเพื่อส่งข้อมูลกลับไปยัง Legacy']  
    
    return pic,text