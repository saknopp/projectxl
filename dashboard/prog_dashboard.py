# -*- coding: utf-8 -*-
from jnj.settings import *
from main.ms_reader_v014 import create_project_dict, close_program
from jnj.ws_dashboard import get_milestones, get_project_status, get_capa_status, get_cat_dash, get_implementation
from xlwings import Workbook, Range
from datetime import datetime

PROGRAM_DASH_TEMPLATE = 'program_dashboard_template.xlsx'

def post_program_dashboard():
    program = create_project_dict(FILE_KEYS, project_list)
    update_program_dashboard(program)
    close_program(program)

def update_program_dashboard(program):
    filename = LOCAL_DASH_PATH + PROGRAM_DASH_TEMPLATE
    print('Excel program template filename: ', filename)
    save_filename = LOCAL_DASH_PATH + 'program_dashboard_' + datetime.today().strftime("%d-%b-%Y_T%H_%M") + '.xlsx'
    wb = Workbook(filename)
    excel = wb.xl_app
    excel.visible = True
    
    ### milestone dashboard
    for file_key in enumerate(FILE_KEYS):
        project = program[file_key[1]]
        row = 4 + 2 * file_key[0]
        due_date_row = Range('team milestones', (row,3), (row,19))
        status_row= Range('team milestones', (row+1,3), (row+1,19))
        milestones = get_milestones(project)
        due_date_row.value = [task.get_due_date() for task in milestones]
        status_row.value = [task.status() for task in milestones]
        
    
    ### CAPA Status
    row=2 
    clear_range = Range('CAPA status', (2,1), (500,11))
    clear_range.value = ''
    for file_key in enumerate(FILE_KEYS):
        project = program[file_key[1]]
        project_status_row = Range('CAPA status', (row,1), (row,10))
        project_status_row.value = get_project_status(project)
        capa_status_row = Range('CAPA status', (row+1,1), (row+1,10))
        capa_status_row.value = get_capa_status(project)
        row = row + 2    
        next_row = Range('CAPA status', (row,1), (row,10))
        IC_cat_dash = get_cat_dash(project, 'IC')
        for task_fields in IC_cat_dash:
            next_row.value = task_fields
            row = row +1 
            next_row = Range('CAPA status', (row,1), (row,11))   
        ### CA / PA dashboard
        for task_fields in get_cat_dash(project, 'CA'):
            next_row.value = task_fields 
            row = row +1 
            next_row = Range('CAPA status', (row,1), (row,11))
        for task_fields in get_cat_dash(project, 'PA'):
            if task_fields != None:
                next_row.value = task_fields
                row = row +1 
                next_row = Range('CAPA status', (row,1), (row,11))       
        ### remediation
        for task_fields in get_cat_dash(project, 'RM'):
            next_row.value = task_fields
            row = row +1
            next_row = Range('CAPA status', (row,1), (row,11))
        row = row +1

        
        
    ### Implementation
    row=2 
    clear_range = Range('implementation', (2,1), (500,11))
    clear_range.value = ''
    for file_key in enumerate(FILE_KEYS):
        project = program[file_key[1]]
        for procedure in get_implementation(project):        
            next_row = Range('implementation', (row,1), (row,10))
            next_row.value = procedure
            row = row + 1    
        
    wb.save(save_filename)
    wb.close()
    excel.quit()      
    
if __name__ == "__main__":
    program = create_project_dict(FILE_KEYS, project_list)
    update_program_dashboard(program)
    close_program(program)

    
