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
    for t in text:
        p = row_content[0].add_paragraph()
        set_text_from_data_bullet(p,t)

    return table

def get_inbound_text(paths,server_name):
    text = []
    text.append('•	ESB ทำการย้ายไฟล์จาก File Server ไปวางที่ SFTP Server')
    for path in paths:
        text.append('      •    '+path)
    text.append('•	SAP PI ทำการหยิบไฟล์ที่ SFTP Server')
    for path in paths:
        text.append('      •    '+path)
    text.append('       เพื่อวางไว้ที่ AL11 บน '+server_name)
    text.append('•	คง AL11 Path, File server และ Path บน File Server ต้นทางไว้ตาม AS-IS')
    return text

def get_outbound_text(paths,server_name):
    text = []
    text.append('•	SAP S/4HANA ทำการวางไฟล์ไว้ที่ AL11')
    text.append('•	SAP PI ทำการหยิบไฟล์ที่ AL11 บน '+server_name+' ไปวางที่ SFTP Server')
    for path in paths:
        text.append('      •    '+path)
    text.append('•	ESB ทำการย้ายไฟล์จาก SFTP Server')
    for path in paths:
        text.append('      •    '+path)
    text.append('       ไปยัง File Server')
    text.append('•	คง AL11 Path, File server และ Path บน File Server ปลายทางไว้ตาม AS-IS')
    return text

def get_pic(data,config):
    pic = None
    text = None
    if "BW/4HANA" in data[col_source_system] or "BW/4HANA" in data[col_target_system]:
        if data[col_inbound]:
            pic = root_images +'/'+ img_file_transfer_sftp +'/bw/'+ "inbound.jpg"
            text = get_inbound_text(data[col_sftp_path],'SAP BW/4HANA')
        else:
            pic = root_images +'/'+ img_file_transfer_sftp +'/bw/'+ "outbound.jpg"
            text = get_outbound_text(data[col_sftp_path],'SAP BW/4HANA')
    elif "S/4HANA" in data[col_source_system] or "S/4HANA" in data[col_target_system]:
        if data[col_inbound]:
            pic = root_images +'/'+ img_file_transfer_sftp +'/hana/'+ "inbound.jpg"
            text = get_inbound_text(data[col_sftp_path],'SAP S/4HANA')
        else:
            pic = root_images +'/'+ img_file_transfer_sftp +'/hana/'+ "outbound.jpg"
            text = get_outbound_text(data[col_sftp_path],'SAP S/4HANA')
    elif "HR" in data[col_source_system] or "HR" in data[col_target_system]:
        if data[col_inbound]:
            pic = root_images +'/'+ img_file_transfer_sftp +'/hr/'+ "inbound.jpg"
            text = get_inbound_text(data[col_sftp_path],'SAP S/4HANA (HR)')
        else:
            pic = root_images +'/'+ img_file_transfer_sftp +'/hr/'+ "outbound.jpg"
            text = get_outbound_text(data[col_sftp_path],'SAP S/4HANA (HR)')
    
    return pic , text
