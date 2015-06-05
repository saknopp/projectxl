# -*- coding: utf-8 -*-
from jnj.settings import *
from main.ms_reader_v014 import return_filename, wrap_Project
from xlwings import Workbook, Sheet, Range
from collections import defaultdict
from datetime import datetime, timedelta


PP_DASH_TEMPLATE = 'team_dashboard_template.xlsx'

def return_file_key(filename):
    for file_key in FILE_KEYS:
        if file_key in filename:
            return file_key

def get_tasks_with_status(project, STATUS):
    task_list = project.task_dict
    """return subset of tasks list with a given status"""
    subset = []
    subset += [task.task for task in task_list if task.status() == STATUS and task.has_subtask() == False]
    return subset
    
def get_pending_tasks(project):
    Task_list = project.Task_list
    """return subset of tasks list for PP dashboard"""
    subset = []
    dash_status = ['<2 WEEKS', '<1 MONTH', 'LATE']
    subset += [task for task in Task_list if task.status() in dash_status and task.has_subtask() == False]
    return subset
    
def get_cat_task_dict(project):
    task_dict = project.task_dict
    cat_task_dict = defaultdict(list)
    for task in task_dict.values():
        key = task.action
        if key != 'N/A':
            if key[:2] in ['IC', 'CA', 'PA', 'RM']:
                if task.has_subtask() == False:
                    cat_task_dict[key].append(task)
                else:
                    cat_task_dict['summary'].append(task)
    return cat_task_dict

def get_cat_dash(project, category):
    cat_task_dict = get_cat_task_dict(project)
    #print(cat_task_dict['summary'])
    ### category == one of 'IC', 'CA', 'PA', or 'RM'
    
    def get_percent_complete(cat_task_dict, next_cat):
            complete = 0
            workload = 0
            for item in cat_task_dict[next_cat]:
                if item.task.PercentComplete > 0 and item.task.Duration != 0:
                    complete += (item.task.Finish - item.task.Start + timedelta(days=1)).days * item.task.PercentComplete / 100
                workload += (item.task.Finish - item.task.Start + timedelta(days=1)).days
            if workload > 0:
                PercentComplete = complete / workload
            else:
                PercentComplete = 'N/A'
            return PercentComplete
            
    cat_dash_list = [cat for cat in list(cat_task_dict.keys()) if cat[:2] == category]
    for cat_increment in range(len(cat_dash_list)):
        next_cat = category + '-' + str(cat_increment + 1)
        next_cat_task = [task for task in cat_task_dict['summary'] if next_cat in task.task.Name][0]
        next_cat_tasklist = [task for task in cat_task_dict[next_cat]]
        cat_increment_summary = next_cat_task.task.Name
        start = next_cat_task.task.Start   
        finish = max([task.task.Finish for task in next_cat_tasklist if isinstance(task.task.Finish, datetime)])
        if [task.task.Deadline for task in next_cat_tasklist if isinstance(task.task.Deadline, datetime)] != []:
            due_date = max([task.task.Deadline for task in next_cat_tasklist if isinstance(task.task.Deadline, datetime)])
        else: 
            due_date = 'N/A'
        PercentComplete = get_percent_complete(cat_task_dict, next_cat)
        status = [task.status() for task in cat_task_dict[next_cat] if task.task.Finish == finish][0]
        
        cat_task_list_open = [task for task in cat_task_dict[next_cat] if task.status() != 'COMPLETE']       
        if cat_task_list_open != []:
            cat_task_next = cat_task_list_open[0]
            for task in cat_task_list_open:
                if task.task.Finish < cat_task_next.task.Finish:
                    cat_task_next = task
        else:
            cat_task_next = [task for task in cat_task_dict[next_cat] if task.task.Finish == finish][0]
        #phase = get_phase(cat_task_next, project)
        if cat_task_next.task.Deadline != 'NA':
            if cat_task_next.task.Deadline.date() < cat_task_next.task.Finish.date():
                next_due_date = cat_task_next.task.Deadline
                mitigation = 'New target date: ' + cat_task_next.task.Finish.strftime("%d-%b-%Y") + '. ' + cat_task_next.recovery_plan
        else:
            next_due_date = cat_task_next.task.Finish
            mitigation = cat_task_next.recovery_plan
        prefix = return_file_key(project.name).split('-')[0]+'-'
        dash_row = [prefix+next_cat, prefix+cat_increment_summary, start, finish, due_date, PercentComplete, status \
        , prefix+cat_task_next.task.Name, next_due_date, cat_task_next.status(), mitigation]
        yield dash_row
    
def get_milestones(project):
    Task_list = project.Task_list
    milestones = []
    for task in Task_list:
        if task.task.Milestone == True:
            milestones.append(task)
    return milestones
        
def get_implementation(project):
        task_dict = project.task_dict
        ws_procedures = task_dict['7.1']
        for item in ws_procedures:
            task = ws_procedures[item]
            #print('Procedure: ', task)
            uniqueid = task.task.UniqueID
            WBS = task.task.WBS
            pred_team = return_file_key(project.name)
            interdependency = 'No'
            proc_owner = return_file_key(project.name)
            doc_number = task.task.Name.split(':')[0]
            title = task.task.Name.split(':')[1]
            change_type = 'Primary'
            ### lookup cn submit task and assign date        
            cn_submit_lookup = project.get_unique_predecessor(WBS)
            
            if cn_submit_lookup != None:
                submit_date = cn_submit_lookup.task.Finish
                #print('Found CN submit.')
            else:
                submit_date = 'TBD'
            ### lookup training complete task and assign date
            current_lookup = cn_submit_lookup

            if current_lookup == None or current_lookup == '>1':
                training_lookup = None
                training_start = 'TBD'
                training_finish = 'TBD'
                #print('Couldnt find submit date')
            else:
                training_lookup = ''
                #print('CN Submit Lookup')  
                #print(type(current_lookup), current_lookup.task.WBS)
                #print('current_lookup.WBS', current_lookup.task.WBS)
            
            while training_lookup == '':
                #print('Training Lookup while')          
                WBS = current_lookup.task.WBS
                training_lookup = project.get_unique_successor(WBS)
                if training_lookup != '>1' and training_lookup != None:
                    if training_lookup.wbs_list[:-1] != [5,2,2,5]:
                        current_lookup = training_lookup                    
                        training_lookup = ''
                        #print('current_lookup.task.WBS', current_lookup.task.WBS)
                    else:
                        training_start = training_lookup.task.Start
                        training_finish = training_lookup.task.Finish
                        #print('Found Training Dates.')
                else: 
                    training_start = 'TBD'
                    training_finish = 'TBD'
                    #print('Training lookup is None or >1')
    
            ###lookup procedure effecive task and assign date
            current_lookup = training_lookup
            #print('Final Training Lookup')
            if current_lookup == None or current_lookup == '>1':
                effective_lookup = None
                effective_date = 'TBD'
            else:
                effective_lookup = ''
    
            while effective_lookup == '':
                #print('effective lookup')                
                #print('procedure: ', task, type(task))                
                #print('current_lookup.WBS', current_lookup)
                WBS = current_lookup.task.WBS
                effective_lookup = project.get_unique_successor(WBS)
                #print(effective_lookup.task.Name)
                if effective_lookup != '>1' and effective_lookup != None:
                    if effective_lookup.wbs_list[:-1] != [5,2,2,6]:
                        current_lookup = effective_lookup                   
                        effective_lookup = ''
                    else:
                        effective_date = effective_lookup.task.Finish
                else: 
                    effective_date = 'TBD'
            task_fields = [uniqueid, pred_team, proc_owner, interdependency, doc_number, title, change_type, submit_date, training_start, \
            training_finish, effective_date]
            yield task_fields
            
def get_capa_status(project):
    file_key=return_file_key(project.name)
    Task_list= project.Task_list
    capa_tasks = []
    for task in Task_list:
        key = task.action
        if key != 'N/A':
            if key[:2] in ['IC', 'CA', 'PA', 'RM']:
                if task.has_subtask() == False:
                    capa_tasks.append(task)
    capa_action = file_key.split('-')[0]+'-'+'CAPA'
    description = file_key.split('-')[0]+'-'+'CAPA'
    current_task = 'N/A'
    task_due_date = 'N/A'
    task_status = 'N/A'
    start = min([task.task.Start for task in capa_tasks if isinstance(task.task.Start, datetime)])
    finish = max([task.task.Finish for task in capa_tasks if isinstance(task.task.Finish, datetime)])
    if [task.task.Deadline for task in capa_tasks if isinstance(task.task.Deadline, datetime)] != []:
        due_date = max([task.task.Deadline for task in capa_tasks if isinstance(task.task.Deadline, datetime)])
    else:
        due_date = 'N/A'
    actual_percent_complete = project.get_percent_complete(capa_tasks)    
    expected_percent_complete = (datetime.today().date() - start.date()) / (finish.date() - start.date())
    if actual_percent_complete < expected_percent_complete - .10:
        CAPA_STATUS = "LATE RISK"
    elif actual_percent_complete < expected_percent_complete - .20:
        CAPA_STATUS = "LATE"
    else:
        CAPA_STATUS = "ON TARGET"
    dash_row = [capa_action, description, start, finish, due_date, actual_percent_complete, CAPA_STATUS, current_task, task_due_date, task_status]           
    return dash_row
    

def get_project_status(project):
    file_key = return_file_key(project.name)    
    Task_list= project.Task_list
    project_tasks = []
    for task in Task_list:    
        if task.has_subtask() == False:
            project_tasks.append(task)
    capa_action = file_key
    description = file_key
    current_task = 'N/A'
    task_due_date = 'N/A'
    task_status = 'N/A'
    start = min([task.task.Start for task in project_tasks if isinstance(task.task.Start, datetime)])
    finish = max([task.task.Finish for task in project_tasks if isinstance(task.task.Finish, datetime)])
    if [task.task.Deadline for task in project_tasks if isinstance(task.task.Deadline, datetime)] != []:
        due_date = max([task.task.Deadline for task in project_tasks if isinstance(task.task.Deadline, datetime)])
    else:
        due_date = 'N/A'
    actual_percent_complete = project.get_percent_complete(project_tasks)    
    expected_percent_complete = (datetime.today().date() - start.date()) / (finish.date() - start.date())
    if actual_percent_complete < expected_percent_complete - .10:
        CAPA_STATUS = "LATE RISK"
    elif actual_percent_complete < expected_percent_complete - .20:
        CAPA_STATUS = "LATE"
    else:
        CAPA_STATUS = "ON TARGET"
    dash_row = [capa_action, description, start, finish, due_date, actual_percent_complete, CAPA_STATUS, current_task, task_due_date, task_status]           
    return dash_row
        
def update_excel_dashboard(project):
    print('Started function.')
    ### filekey requires at least first two characters filename 
    file_key = return_file_key(project.name)
    dash_filename =  return_filename(file_key, dash_list)       
    ### dispatch project class object, importing data into Project Class    
    if dash_filename != None:
        filename = LOCAL_DASH_PATH + dash_filename
        save_filename = LOCAL_DASH_PATH + file_key + '-dashboard-' + datetime.today().strftime("%d-%b-%Y_T%H_%M") + '.xlsx'
    else:
        filename = LOCAL_DASH_PATH + PP_DASH_TEMPLATE
        save_filename = LOCAL_DASH_PATH + file_key + '-dashboard-' + datetime.today().strftime("%d-%b-%Y_T%H_%M") + '.xlsx'  
    wb = Workbook(filename)
    excel = wb.xl_app
    excel.visible = True
    pending_tasks = Sheet('pending_tasks')
    project_plan = Sheet('project_plan')

    sheet_dict = {'project_plan': (project_plan, project.Task_list), 'pending_tasks': (pending_tasks, get_pending_tasks(project))}
    cat_task_dict = get_cat_task_dict(project)    
    for sheet_name in list(sheet_dict.keys())[:2]:
        row =2
        next_row = Range(sheet_dict[sheet_name][0].name, (row,1), (row,11))
        clear_range = Range(sheet_dict[sheet_name][0].name, (2,1), (500,11))
        clear_range.value = ''
        for item in sheet_dict[sheet_name][1]:
            task = item.task
            task_fields = [task.WBS, item.action, project.get_phase(item), task.Name, task.Start, task.Finish, task.Deadline, task.PercentComplete, \
            item.status(), item.status_override, item.recovery_plan]
            next_row.value = task_fields
            row = row +1
            next_row = Range(sheet_dict[sheet_name][0].name, (row,1), (row,9))
    print('Updated project plan and pending tasks.')
    ### milestone dashboard
    row = 5
    next_row = Range('milestones', (row,2), (row,4))
    for item in get_milestones(project):
        due_date = "Not assigned"
        task = item.task
        if task.Deadline != 'NA':
            if task.Deadline.date() < task.Finish.date() and task.PercentComplete != 100:
                due_date = task.Deadline
                mitigation = 'New target date: ' + task.Finish.strftime("%d-%b-%Y") + '. ' + item.recovery_plan
            else:
                due_date = task.Finish
                mitigation = item.recovery_plan
        else:
            due_date = task.Finish
            mitigation = item.recovery_plan
        task_fields = [task.Name, due_date, item.status(), mitigation]
        next_row.value = task_fields
        row = row +1
        next_row = Range('milestones', (row,2), (row,4))     
    print('Updated milestones.')
    row = 2    
    clear_range = Range('CAPA status', (2,1), (500,11))
    clear_range.value = ''
    ### project
    project_status_row = Range('CAPA status', (row,1), (row,10))
    task_fields = get_project_status(project)
    project_status_row.value = task_fields
    project_status = Range('milestones', (1,4), (1,4))
    project_status.value = task_fields[6]

    capa_status_row = Range('CAPA status', (row+1,1), (row+1,10))
    task_fields = get_capa_status(project)
    capa_status_row.value = task_fields
    capa_status = Range('milestones', (2,4), (2,4))
    capa_status.value = task_fields[6]
    row = row + 2   
    
    ### interim controls
    next_row = Range('CAPA status', (row,1), (row,11))
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
    print('Updated CAPA status.')




    
    
    #### implementation
    row = 2
    clear_range = Range('implementation', (2,1), (500,11))
    clear_range.value = ''
    next_row = Range('implementation', (row,1), (row,11))
    for procedure in get_implementation(project):
        next_row.value = procedure
        row = row +1
        next_row = Range('implementation', (row,1), (row,11))
    print('Updated implementation.')
    wb.save(save_filename)
    #print(save_filename)
    wb.close()
    #excel.quit()          

def update_dashboard_by_key(file_key):
    filename = LOCAL_MPP_PATH + return_filename(file_key, project_list)
    project = wrap_Project(filename)
    update_excel_dashboard(project)    
    project.close_project()        

def update_all_dashboards():
    for file_key in FILE_KEYS:
        print(file_key)
        update_dashboard_by_key(file_key)
                    
if __name__ == "__main__":
    pass

    



    
    
    
    
    
    
    
    
 
