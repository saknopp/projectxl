from jnj.settings import *
import win32com.client as win
from datetime import datetime
from collections import defaultdict, namedtuple
import re
import Pyro4
from xlwings._xlwindows import _com_time_to_datetime as convert_date

#pjApp = win.Dispatch("MSProject.Application")
pjApp = win.gencache.EnsureDispatch("MSProject.Application")
pjApp.Visible  = 0

def return_filename(file_key, project_list):
    for filename in project_list:
        if file_key in filename:
            return filename           
     
def open_mpp(filename):
        """
        Returns a win32com client instance of MSProject ActiveProject
        """
        ###pjApp = win.Dispatch("MSProject.Application")
        ### pjApp.visible  = False
        print(filename)        
        pjApp.FileOpen(filename)
        pjApp.OutlineShowAllTasks()
        return pjApp.ActiveProject

def create_program_dict(FILE_KEYS, project_list):
    program_dict = {}
    for file_key in FILE_KEYS:
        filename = LOCAL_MPP_PATH + return_filename(file_key, project_list)  
        mpp = open_mpp(filename)  
        project = Project(mpp)
        program_dict[file_key] = project
    return program_dict
        
def close_program(program):
    for file_key in program:
        pjApp.FileSave()        
        pjApp.FileClose()        
    pjApp.Quit()
        
def wrap_Project(filename):
    """
    Returns a wrapped Project Class instance for an MSProject file
    """
    mpp = open_mpp(filename)    
    project = Project(mpp)
    return project
    
def wrap_create_program_dict():
    program_dict = create_program_dict(FILE_KEYS, project_list)
    return program_dict
    
    

TaskState = namedtuple('TaskState', ['wbs_list', 'action', 'name', 'has_subtask', 'start', 'finish', 'due_date', 'percent_complete', 'status', 'status_override', 'recovery_plan'])

class Task(defaultdict):
    """
    An MS project task class dictionary object derived from win32com.client.Dispatch Item object
    """    
    def get_due_date(self):
        if self.task.Deadline != 'NA':
            if self.task.Deadline.date() < self.task.Finish.date() and self.task.PercentComplete != 100:
                due_date = self.task.Deadline
            else:
                due_date = self.task.Finish
        else:
            due_date = self.task.Finish
        return due_date
        
    def status(self):
        if self.task.PercentComplete == 100:
            return "COMPLETE"           
        elif self.task.Finish.date()  < datetime.today().date() or (self.task.Deadline != 'NA' and self.task.Deadline.date() < datetime.today().date()):
            if self.recovery_plan == '':
                self.recovery_plan = 'Recovery Plan needed.'
            if self.status_override != 'N/A':
                self.staus_overide = 'N/A'
            return "LATE"
        elif self.status_override == "ON TARGET":
            return "ON TARGET"
        elif (self.task.Deadline != 'NA' and self.task.Finish.date() > self.task.Deadline.date()) or self.status_override == "AT RISK":
            if self.recovery_plan == '':
                self.recovery_plan = 'Status Update needed.'        
            return "LATE RISK"
        elif (self.task.Finish.date() - datetime.today().date()).days < 14:
            if self.recovery_plan == '':
                self.recovery_plan = 'Status Update needed.'        
            return "<2 WEEKS"
        elif (self.task.Finish.date() - datetime.today().date()).days < 30:
            if self.recovery_plan in ['Recovery Plan needed.', 'Status Update needed.']:
                self.recovery_plan = ''            
            return "<1 MONTH"
        elif (self.task.Finish.date() - datetime.today().date()).days < 90:
            if self.recovery_plan in ['Recovery Plan needed.', 'Status Update needed.']:
                self.recovery_plan = ''            
            return "<3 MONTH"        
        
        
        else:
            if self.recovery_plan in ['Recovery Plan needed.', 'Status Update needed.']:
                self.recovery_plan = ''   
            return "ON TARGET"
    
    def has_subtask(self):
        return self != {}
        
    def walk_tasks(self):
        for level, subtask_dict in self.items():
            yield subtask_dict.task
     
    def __init__(self, task):
        defaultdict.__init__(self)
        self.task = task
        self.wbs_list = list(map((lambda x: int(x)), self.task.WBS.split('.')))
        self.status_override = self.task.Text5
        if self.status_override == '':
            self.status_override = 'N/A'
        self.recovery_plan = self.task.Text6
        self.action = 'N/A'
        if ':' in task.Name:
            test_action = self.task.Name.split(':')[0]
            if test_action[:2] in ['IC', 'CA', 'PA', 'RM']:
                self.action = test_action
        self.task_state = TaskState(wbs_list=self.wbs_list, action=self.action, name=self.task.Name, \
        has_subtask = self.has_subtask(), start=convert_date(self.task.Start), finish=convert_date(self.task.Finish), due_date=convert_date(self.get_due_date()), \
        percent_complete=self.task.PercentComplete, status=self.status(), status_override=self.status_override, \
        recovery_plan=self.recovery_plan)

    def __str__(self):
        output = '{0}: '.format(self.task.WBS) + '{0}\n'.format(self.task.Name)
        for key, value in self.items():
            output += self[key].__str__()
        return output
        
    def __getstate__(self):
        
        result = self.task_state
        return result
        
class Project(defaultdict):
    """
    A project class to encapsulate the attributes of the MS Project Dispatch 
    instance in python friendly data structures.
    """            
                 
    def get_task(self, wbs_list):
        if wbs_list == []:
            return_dict= self
        for level in enumerate(wbs_list):
            if level[0] == 0:
                return_dict = self[level[1]]
            elif level[0] > 0 :            
                return_dict = return_dict[level[1]]
        return return_dict    
    
    def get_unique_predecessor(self, WBS):
        if self.task_dict[WBS].task.Predecessors != '':
            if len(self.task_dict[WBS].task.Predecessors.split(',')) == 1:        
                predecessor_reference =  int(''.join(filter(lambda x : x.isdigit(), self.task_dict[WBS].task.Predecessors)))     
                unique_predecessor_task = self.task_list[predecessor_reference - 1] 
                unique_predecessor = self.task_dict[unique_predecessor_task.WBS]
            elif len(self.task_dict[WBS].task.Predecessors.split(',')) > 1:
                unique_predecessor = '>1'
        else:
            unique_predecessor = None
        return unique_predecessor
    
    def get_unique_successor(self, WBS):
        if self.task_dict[WBS].task.Successors != '':
            #print('Successors: ', self.task_dict[WBS].task.Successors)
            successors = self.task_dict[WBS].task.Successors.split(',')
            successor = [s for s in successors if 'FF' not in s]
            #print('successors: ', successor)
            if len(successor) == 1:              
                successor_reference = int(re.match('\d+', successor[0]).group())
                #print('successor_reference: ', successor_reference)
                unique_successor_task = self.task_list[successor_reference-1]
                #print('unique_successor_task: ', unique_successor_task.WBS)
                unique_successor = self.task_dict[unique_successor_task.WBS]
                #print('unique_successor: ', unique_successor.wbs_list)
            elif len(successor) > 1:
                unique_successor = '>1'
                successor_reference = None
                successors = None
        else: 
            unique_successor = None
            successor_reference = None
            successors = None
        #print('unique_successor: ', unique_successor)
        return unique_successor
        
    def get_phase(self, task):
        WBS = str(task.wbs_list[0])
        return self.task_dict[WBS].task.Name     
    
    def get_percent_complete(self, Task_list):
        complete = 0
        workload = 0
        for item in Task_list:
            if item.task.PercentComplete > 0:
                complete += (item.task.Finish - item.task.Start).days * item.task.PercentComplete / 100
            workload += (item.task.Finish - item.task.Start).days
        if workload > 0:
            PercentComplete = complete / workload
        else:
            PercentComplete = 'N/A'
        return PercentComplete
    
    def __str__(self):
        output = ''
        for key, value in self.items():
            output += self[key].__str__()
        return output

    def __init__(self, mpp):
        defaultdict.__init__(self)
        self.task_dict = {}
        self.Task_list = []
        self.name = mpp.Name
        self.tasks = mpp.Tasks
        self.task_list = [task for task in self.tasks]
        self.task_state_dict = {}
        
        def lookup_dict(self, wbs_list):
            if wbs_list == []:
                return_dict= self
            for level in enumerate(wbs_list):
                if level[0] == 0:
                    return_dict = self[level[1]]
                elif level[0] > 0 :            
                    return_dict = return_dict[level[1]]
            return return_dict         
                
        for task in self.task_list:
            if task != None:
                wbs_list = list(map((lambda x: int(x)), task.WBS.split('.')))        
                update_dict = lookup_dict(self, wbs_list[:-1])
                new_Task = Task(task)
                update_dict[wbs_list[-1]] = new_Task
                self.task_dict[task.WBS]= new_Task
                self.task_state_dict[task.WBS] = new_Task.task_state
                self.Task_list.append(new_Task)
    
    def close_project(self):
        pjApp.FileSave()        
        pjApp.FileClose()
        
    def __getstate__(self):
        result = self.task_state_dict
        return result
        
    


             
if __name__ == "__main__":
    program_dict = create_project_dict(FILE_KEYS, project_list)
    
"""
### Need project task status and recovery plan fields to read and eventually write back into file
###      include logic override AT RISK by lead -- flag on impoort
### 
### Powerpoint module to populate Team dashboard slides and Overview slides    
"""