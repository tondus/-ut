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

def get_pic(data,config):
    pic = None
    text = [] 

    if data[col_inbound]:
        pic =   root_images +'/'+ img_file_transfer +'/'+ "inbound.png"
        text.append(config['default']['file_transfer_inbound_line1'])
        text.append(config['default']['file_transfer_inbound_line2'])
        text.append(config['default']['file_transfer_inbound_line3'])
    else : 
        pic =   root_images +'/'+ img_file_transfer +'/'+ "outbound.png"
        text.append(config['default']['file_transfer_outbound_line1'])
        text.append(config['default']['file_transfer_outbound_line2'])
        text.append(config['default']['file_transfer_outbound_line3'])
        text.append(config['default']['file_transfer_outbound_line4'])
    return pic , text
