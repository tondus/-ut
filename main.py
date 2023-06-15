import sys 
import pandas as pd
import openpyxl
from slugify import slugify
import os
import pprint as pp
from copy import copy, deepcopy
from openpyxl.utils import rows_from_range
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from distutils.dir_util import copy_tree
from pathlib import Path
import subprocess
from os import walk
import math
import shutil
import pydash
import ctp_template as ctp
from pathlib import Path
import configparser


col_ricefw_id = 'ricefw_id'
col_ricefw_name = 'ricefw_name'
col_document_owner = 'document_owner'
col_date = 'date'
col_tester =  'tester'
col_program_type =  'program_type'
col_middleware =  'middleware'
col_condition_no =  'condition_no'
col_condition_desc =  'condition_desc'
col_expected_result =  'expected_result'
col_fig_title1 =  'fig_title1'
col_fig_title2 =  'fig_title2'
col_fig_title3 =  'fig_title3'
col_fig_title4 =  'fig_title4'
col_fig_title5 =  'fig_title5'
col_pass_fail =  'pass_fail'
col_actual_result =  'actual_result'
col_payload =  'payload'

text_file = ['.xml','.json','.txt','.dat']
def replace_from_onedrive(p):
    for user_dir in os.listdir(p):
        local = 'onedrive/'+user_dir
        if not os.path.exists(local):
            os.mkdir(local)
        master_file = p+'\\'+user_dir+'\\'+'OREO.xlsx'
        master_images = p+'\\'+user_dir+'\\'+'images'
        dest_file = local + '/' +'OREO.xlsx'
        dest_images = local + '/' +'images'
        if os.path.exists(master_file):
             shutil.copy(master_file, dest_file)
        if os.path.exists(master_images):
            if os.path.exists(dest_images):
                 shutil.rmtree(dest_images) 
            shutil.copytree(master_images, dest_images)

def column_to_path(columns):
    result = []
    path = []
    for c in columns:
        path.append(c)
        result.append(copy(path))
    return result

def get_new_step(row):
    return {'expected_result':row[col_expected_result],
                           'pass_fail':row[col_pass_fail],
                           'fig_title1':row[col_fig_title1],
                           'fig_title2':row[col_fig_title2],
                           'fig_title3':row[col_fig_title3],
                           'fig_title4':row[col_fig_title4],
                           'fig_title5':row[col_fig_title5],}

def create_structure(df,columns):
    result = {}
    for idx,row in df.iterrows():
        if row[col_ricefw_id] in result:
            if row[col_condition_no] in result[row[col_ricefw_id]]['test_cases']:
                steps = result[row[col_ricefw_id]]['test_cases'][row[col_condition_no]]['steps']
                steps.append(get_new_step(row))
            else:
                test_cases = result[row[col_ricefw_id]]['test_cases']
                test_cases[row[col_condition_no]] = {'condition_no':row[col_condition_no],
                                                    'condition_desc':row[col_condition_desc],
                                                    'steps':[get_new_step(row)]}
        else:
            result[row[col_ricefw_id]] = {  "ricefw_id": row[col_ricefw_id],
                                            "ricefw_name": row[col_ricefw_name],
                                            "document_owner":row[col_document_owner],
                                            "date":row[col_date],
                                            "program_type":row[col_program_type],
                                            "middleware":row[col_middleware],
                                            "test_cases": {row[col_condition_no]: {'condition_no':row[col_condition_no],
                                                                                    'condition_desc':row[col_condition_desc],
                                                                                    'steps':[get_new_step(row)]}}}
    return result
          

def fill_up_dataframe(df,columns):
    result = {}
    exception_col = ['fig_title1','fig_title2','fig_title3','fig_title4','fig_title5','pass_fail']
    lastest = {}
    for idx , row in df.iterrows():
        for col in columns:
            if col in result:
                values = result[col]
                if not pd.isna(row[col]):
                    lastest[col] = row[col]
                if col not in exception_col:
                    values.append(lastest[col])
                else:
                    values.append(row[col])
                result[col] = values
            else:
                result[col] = [row[col]]
                lastest[col] = row[col]
    
    df_master = pd.DataFrame()
    for key in result.keys():
        df_master[key] = result[key]
    return df_master

def generate_images_folder(program,path):
    for key in program.keys():
        prog = program[key]
        ricefw_id = prog['ricefw_id']
        figure_index = 1
        for t in prog['test_cases'].keys():
            test_cases =  prog['test_cases'][t]
            for step in test_cases['steps']:
                figure = 'Figure '+str(figure_index)
                figure_index += 1
                full_path = path +ricefw_id+ "/" +figure
                if not os.path.isdir(full_path):
                    os.makedirs(full_path)

def adjust_image_size(image_dir,config):
    resize = " "+str(config['DEFAULT']['ImageSize']) +"% "
    # resize = " 80% "
    p = 'source/image_magick/'
    if os.path.isdir(p):
        shutil.rmtree(p)
    for path, subdirs, files in os.walk(image_dir):
        for name in files:
            _filename, extension = os.path.splitext(name)
            if extension.lower() not in text_file:
                program_path = path.replace(image_dir,"")
                program_path = program_path.replace('\\','/')
                magick_path = 'source/image_magick/'+program_path+'/'
                filename = os.path.join(path, str(name))
                filename = filename.replace('\\','/')
                name = os.path.basename(filename)
                if os.path.isfile(filename):
                    Path(magick_path).mkdir(parents=True, exist_ok=True)
                    subprocess.call(r'magick convert '+ '"'+filename +'"'+ " -resize "+resize +'"' +magick_path + name+'"', shell=True)

def copy_images_to_result(root_path,user_dir):
    local_images_path = 'onedrive/'+user_dir+'/images'
    if os.path.exists(local_images_path):
        result_path = root_path+'\\'+user_dir +'\\'+'generated_results'
        result_images_path = root_path+'\\'+user_dir +'\\'+'generated_results' +'\\'+'images'
        if not os.path.exists(result_path):
            os.mkdir(result_path)
        if os.path.exists(result_images_path):
            shutil.rmtree(result_images_path)
        shutil.copytree(local_images_path, result_images_path)
    
def main() -> int:
    config = configparser.ConfigParser()
    config.read('config.ini')
    df = pd.read_excel('source/OREO.xlsx')
    column_headers = list(df.columns.values)
    df_master = fill_up_dataframe(df,column_headers)
    data = create_structure(df_master,column_headers)
    generate_images_folder(data,'source/images/')
    adjust_image_size('source/images/',config)
    ctp.create_document(data,config)
                    
if __name__ == '__main__':
    sys.exit(main())