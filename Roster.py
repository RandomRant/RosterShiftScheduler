#This is an duty allocation optimization model for the KKH DPM CC Team, built by Suraj Kamath, kamath.s.suraj@gmail.com
#its objective is to assign duties as fairly as possible across Clinical Counsellors while fulfilling the operational requirements of the team

import os
import pandas as pd
import pulp
import xlwings as xw
from tkinter import filedialog
import tkinter as tk


TOTAL_PERIODS_NUM = 10
TOTAL_DUTIES_NUM = 4
#Periods are defined as Mon AM, Mon PM, Tue AM ... Fri PM
#Duties currently are Child CC inpatient, Child CC BL/outpatient, Women CC inpatient 



def model_problem(week_to_optimize):

    clinic_constraint=int(sheet.range(6,39).value)
    allow_scr_cc_repeat=int(sheet.range(7,39).value)
    max_weekly_duties=int(sheet.range(8,39).value)
    max_weekly_oncalls=int(sheet.range(9,39).value)
    max_weekly_child_oncalls=int(sheet.range(10,39).value)
    c_fairness_weight=int(sheet.range(11,39).value)
    w_fairness_weight=int(sheet.range(12,39).value)
    s_fairness_weight=int(sheet.range(13,39).value)
    day_oncall_samecc_weight=int(sheet.range(14,39).value)
    seconds_to_optimize=int(sheet.range(15,39).value)
    
    #row offsets to apply for each week. this wont change unless number of duties change of number of ccs exceeds 15
    ass_offset=(week_to_optimize-1)*11
    avl_offset=(week_to_optimize-1)*18
    total_num_counselors=len([x for x in sheet.range((91,2),(105,2)).value if x is not None]) #gets #of counsellors 

    #main program, runs and returns worker data and solved Lp problem
    
    #delete existing data in week duty allocs. Later might include warning on delete 
    for period in range(TOTAL_PERIODS_NUM):
        sheet.range((ass_offset+8,4+period*2),(ass_offset+8+TOTAL_DUTIES_NUM,4+period*2)).options(transpose=True).value=""

    #below code reads each row of CC Availability from below the roster sheet
    workerdf = pd.DataFrame(data=sheet.range((avl_offset+91,1),(avl_offset+101,23)).value)
    
    #below code grabs clinic assignment data
    clinicdf= pd.DataFrame(data=sheet.range((ass_offset+4,3),(ass_offset+4+2,3+20)).value).drop(labels=[0,2,4,6,8,10,12,14,16,18,20],axis=1).reindex(axis=1,method='pad')
    

    #Read Data into a dictionary to help setup the optimization problem
    workers_data = {} #initializes a dictionary called workers_data
    for iteration in workerdf.iterrows(): 
        row = iteration[1] #gets the data of each row as a series object
        name = row[1] #gets the worker name from the series object
        workers_data[name] = {} #initializes a nested dictionary under workers_data with name:P{}
        
        workers_data[name]["period_avail"] = [] #initializes a list, nested dict structure now {name:{skill_level:row[1]value,period_avail:[]}}
        workers_data[name]["clinic_assign"] = []
        workers_data[name]["existing_assign"]=[]
        workers_data[name]["change_assign"]=[]
       
        
        #this code checks if any clinic has been assigned for the week for each CC and puts those assignment in a list called clinic_assign
        for period in range(TOTAL_PERIODS_NUM):            
            clinic=0
            for c_row in range(3):
                clinic +=str(clinicdf.iloc[c_row,period]).find(name)
            if clinic==-3: 
                workers_data[name]["clinic_assign"].append(0)
            else:
                workers_data[name]["clinic_assign"].append(1)

        #Read in existing data as already rostered information
        for period in range(TOTAL_PERIODS_NUM):
            for duty in range(TOTAL_DUTIES_NUM):
                if sheet.range(ass_offset+8+duty,4+period*2).value==name:
                    workers_data[name]["existing_assign"].append(1)
                else:
                    workers_data[name]["existing_assign"].append(0)

    #delete existing data in week duty allocs
    for period in range(TOTAL_PERIODS_NUM):
       sheet.range((ass_offset+8,4+period*2),(ass_offset+8+TOTAL_DUTIES_NUM,4+period*2)).options(transpose=True).value=""                
                    
    #get worker availability data
    workerdf = pd.DataFrame(data=sheet.range((avl_offset+91,1),(avl_offset+101,23)).value)
    for iteration in workerdf.iterrows(): 
        row = iteration[1] #gets the data of each row as a series object
        name = row[1] #gets the worker name from the series object  
        for period in range(TOTAL_PERIODS_NUM): 
                if row[period*2+3]!=name:
                    workers_data[name]["period_avail"].append(0)
                else:
                    workers_data[name]["period_avail"].append(1)
            
    #problem is defined as minimizing the objective function which is total worked periods
    problem = pulp.LpProblem("ScheduleWorkers", pulp.LpMinimize)
    tvars=[]
    cvars=[]
    workerid = 0
    for worker in workers_data.keys():
        workerstr = str(workerid)
        periodid = 0
        dutyid=0

        workers_data[worker]["assigned_period_duty"] = []
        
        #define 1 variables for each worker_period_duty to optimize. each are binary values.the lists in worker_data hold the variable names and pointers
        for period_avl in workers_data[worker]["period_avail"]: 
            periodstr = str(periodid)
            for dutyid in range(TOTAL_DUTIES_NUM):
                dutystr=str(dutyid)
                existing_assign=workers_data[worker]["existing_assign"][periodid*TOTAL_DUTIES_NUM+dutyid]
                # worked periods: worker W works in period P
                workers_data[worker]["assigned_period_duty"].append(
                    pulp.LpVariable("x_{}_{}_{}".format(workerstr, periodstr,dutystr), cat=pulp.LpInteger,lowBound=0, upBound=period_avl)) #upBound ensures that for unavailable period, the upbound is zero, so nothing can be assigned
                workers_data[worker]["change_assign"].append(pulp.LpVariable("ch_{}_{}_{}".format(workerstr, periodstr,dutystr), cat=pulp.LpInteger,lowBound=0, upBound=1)) #define chvars to apply penalties for changes relative to existing assignments
                dutyid +=1
            periodid += 1
        workerid += 1
        
        #create tvars to assign penalty for unfair allocation of duties across cc
        dutyid=0
        for dutyid in range(TOTAL_DUTIES_NUM):
            dutystr=str(dutyid)
            tvars.append(pulp.LpVariable("t_{}_{}".format(workerstr,dutystr), cat=pulp.LpInteger))
        
        #create cvars to assign penalty for duty0,1,2 allocations across diff ccs for 2 periods in same days
        periodid=0
        for period in range(0,TOTAL_PERIODS_NUM,2):
            periodstr=str(period)
            for dutyid in range(0,TOTAL_DUTIES_NUM-1):
                dutystr=str(dutyid)
                cvars.append(pulp.LpVariable("c_{}_{}_{}".format(workerstr,periodstr,dutystr), cat=pulp.LpInteger,lowBound=0, upBound=1))
 
   
    # Create objective function (amount of turns worked). The Objective fn is the total number of worked periods
    objective_function = None 
    tvars_c = None  #fairness variables for child duties
    tvars_w = None #fairness for womens duties
    tvars_s = None #fairness for screening duties
    chvars = None  #varible to impose penalties for changes to existing roster assignments(reduce changes when reoptimizing)

    workerid=0
    for worker in workers_data.keys():
        for dutyid in range(0,2): 
            tvars_c += tvars[dutyid+workerid*TOTAL_DUTIES_NUM] 
        tvars_w += tvars[2+workerid*TOTAL_DUTIES_NUM] 
        tvars_s += tvars[3+workerid*TOTAL_DUTIES_NUM]     

        for period in range(TOTAL_PERIODS_NUM):
            for dutyid in range(TOTAL_DUTIES_NUM):
                if workers_data[worker]["existing_assign"][dutyid+period*TOTAL_DUTIES_NUM]==1:    
                        chvars+=workers_data[worker]["change_assign"][dutyid+period*TOTAL_DUTIES_NUM]
        
        workerid+=1
        
    objective_function = c_fairness_weight*tvars_c + w_fairness_weight*tvars_w +s_fairness_weight*tvars_s+day_oncall_samecc_weight*cvars 
    if not chvars== None: objective_function += 3*chvars
    problem += objective_function
    
    #creates a constraint for each CC and period  CC can only be assigned 1 duty
    if allow_scr_cc_repeat==0:
        for period in range(TOTAL_PERIODS_NUM):

            for worker in workers_data.keys():
                work_period_duty_sum=None
                for duty in range(TOTAL_DUTIES_NUM):
                    work_period_duty_sum += workers_data[worker]["assigned_period_duty"][period*TOTAL_DUTIES_NUM+duty] 
                problem += work_period_duty_sum  <= 1
    else:
        for period in range(TOTAL_PERIODS_NUM):

            for worker in workers_data.keys():
                work_period_duty_sum=None
                for duty in range(TOTAL_DUTIES_NUM-1):
                    work_period_duty_sum += workers_data[worker]["assigned_period_duty"][period*TOTAL_DUTIES_NUM+duty] 
                problem += work_period_duty_sum  <= 1

    #creates a constraint for each CC cannot have more than 6 duties
    for worker in workers_data.keys():
        work_period_duty_sum=None
        for x in range(TOTAL_PERIODS_NUM*TOTAL_DUTIES_NUM):
                work_period_duty_sum += workers_data[worker]["assigned_period_duty"][x] 
        problem += work_period_duty_sum  <= max_weekly_duties

    #creates a constraint for each CC cannot have more than 4 child on calls
    for worker in workers_data.keys():
        work_period_duty_sum=None
        for period in range(TOTAL_PERIODS_NUM):
                work_period_duty_sum += workers_data[worker]["assigned_period_duty"][0+period*TOTAL_DUTIES_NUM] + workers_data[worker]["assigned_period_duty"][1+period*TOTAL_DUTIES_NUM] 
        problem += work_period_duty_sum  <= max_weekly_child_oncalls

    #creates a constraint for each CC cannot have more than 5 child+women on calls
    for worker in workers_data.keys():
        work_period_duty_sum=None
        for period in range(TOTAL_PERIODS_NUM):
                work_period_duty_sum += workers_data[worker]["assigned_period_duty"][0+period*TOTAL_DUTIES_NUM]+workers_data[worker]["assigned_period_duty"][1+period*TOTAL_DUTIES_NUM] +workers_data[worker]["assigned_period_duty"][2+period*TOTAL_DUTIES_NUM] 
        problem += work_period_duty_sum  <= max_weekly_oncalls

 

    # Each period_Duty must have one and only one assignee
    for period in range(TOTAL_PERIODS_NUM):
        for duty in range(TOTAL_DUTIES_NUM):
            work_period_duty_sum=None
            for worker in workers_data.keys():
                 work_period_duty_sum += workers_data[worker]["assigned_period_duty"][duty+(period*TOTAL_DUTIES_NUM)] 
            problem += work_period_duty_sum  == 1
            
   
    #a cc cannot be assigned a child CC or BL "(duty0 and duty1)" on the day they are assigned clinic
    if clinic_constraint==0:
        workerid=0
        for worker in workers_data.keys():
            work_period_duty_sum=None
            for period in range(0,TOTAL_PERIODS_NUM,2):
                work_period_duty_sum = workers_data[worker]["assigned_period_duty"][0+(period+1)*TOTAL_DUTIES_NUM]+workers_data[worker]["assigned_period_duty"][1+(period+1)*TOTAL_DUTIES_NUM]
                #print(period," ",work_period_duty_sum)
                if  workers_data[worker]["clinic_assign"][period]==1: problem += work_period_duty_sum <= 0
            for period in range(1,TOTAL_PERIODS_NUM,2):
                work_period_duty_sum = workers_data[worker]["assigned_period_duty"][0+(period-1)*TOTAL_DUTIES_NUM]+workers_data[worker]["assigned_period_duty"][1+(period-1)*TOTAL_DUTIES_NUM] 
                #print(period," ",work_period_duty_sum)
                if  workers_data[worker]["clinic_assign"][period]==1: problem += work_period_duty_sum <= 0


    #Constraint to impose penalty if on a single day duty0,1,2 both periods are not assigned to the same CC
    workerid=0
    for worker in workers_data.keys():
        work_duty_sum=None
        for period in range(0,TOTAL_PERIODS_NUM,2):
            for dutyid in range(3):
                work_duty_sum = workers_data[worker]["assigned_period_duty"][dutyid+period*TOTAL_DUTIES_NUM] - workers_data[worker]["assigned_period_duty"][dutyid+(period+1)*TOTAL_DUTIES_NUM]
                #print("w-",work_duty_sum," ", cvars[int(dutyid+(period)*(TOTAL_DUTIES_NUM-1)/2)+(15*workerid)]," ",int(dutyid+(period)*(TOTAL_DUTIES_NUM-1)/2)+(15*workerid))
                problem += work_duty_sum <= cvars[int(dutyid+(period)*(TOTAL_DUTIES_NUM-1)/2)+(15*workerid)]
                problem += -work_duty_sum  <= cvars[int(dutyid+(period)*(TOTAL_DUTIES_NUM-1)/2)+(15*workerid)]
        workerid+=1
   
        

    #Constraint for abs of difference for fair allocation  
    for dutyid in range(0,TOTAL_DUTIES_NUM):
        workerid=0
        #print(dutyid,workerid)
        allworker_duty_sum=None
        for worker in workers_data.keys():
            for period in range(TOTAL_PERIODS_NUM):
                allworker_duty_sum += workers_data[worker]["assigned_period_duty"][dutyid+period*TOTAL_DUTIES_NUM]
        allworker_duty_avg=allworker_duty_sum/10
        for worker in workers_data.keys():
            work_duty_sum=None
            for period in range(TOTAL_PERIODS_NUM):
                work_duty_sum += workers_data[worker]["assigned_period_duty"][dutyid+period*TOTAL_DUTIES_NUM]
            #print("w-",work_duty0_sum," ", tvars[0+workerid*TOTAL_DUTIES_NUM])
            problem += work_duty_sum - allworker_duty_avg <= tvars[dutyid+workerid*TOTAL_DUTIES_NUM]
            problem += -(work_duty_sum - allworker_duty_avg) <= tvars[dutyid+workerid*TOTAL_DUTIES_NUM]
            workerid +=1

    #Constraint to impose penalty on changes made relative to existing assignments
    workerid=0
    for worker in workers_data.keys():
        for period in range(TOTAL_PERIODS_NUM):
            for dutyid in range(TOTAL_DUTIES_NUM):
                change_duty_sum=None
                if workers_data[worker]["existing_assign"][dutyid+period*TOTAL_DUTIES_NUM]==1:
                    change_duty_sum = workers_data[worker]["assigned_period_duty"][dutyid+period*TOTAL_DUTIES_NUM] - workers_data[worker]["existing_assign"][dutyid+period*TOTAL_DUTIES_NUM]
                    problem += change_duty_sum <= workers_data[worker]["change_assign"][int(dutyid+period*(TOTAL_DUTIES_NUM))]
                    problem += -change_duty_sum  <= workers_data[worker]["change_assign"][int(dutyid+period*(TOTAL_DUTIES_NUM))]
        workerid+=1

    
    #Execute Solver with error handling
    solver = pulp.getSolver('GLPK_CMD',timeLimit=seconds_to_optimize,mip=True,options=['--mipgap','0.2'])
    try:
        x=problem.solve(solver)
    except Exception as e:
        print("Can't solve problem: {}".format(e))


    #check the worked periods list for each worker. if the period value is 1 i.e the worker is allocated that shift, then add it to the schedule
    df_out=pd.DataFrame(index=["Duty 1","Duty 2","Duty 3","Duty 4"], columns=["MON AM","MON PM","TUE AM","TUE PM","WED AM","WED PM","THU AM","THU PM","FRI AM","FRI PM"])

    for period in range(TOTAL_PERIODS_NUM):
        for duty in range(TOTAL_DUTIES_NUM):
            for worker in workers_data.keys():
                if x==1:
                    if workers_data[worker]["assigned_period_duty"][duty+(period*TOTAL_DUTIES_NUM)].varValue==1: 
                        df_out.iloc[duty,period]=worker
                else:
                    if workers_data[worker]["existing_assign"][duty+(period*TOTAL_DUTIES_NUM)]==1: 
                        df_out.iloc[duty,period]=worker

    for period in range(TOTAL_PERIODS_NUM):
        sheet.range((ass_offset+8,4+period*2),(ass_offset+8+TOTAL_DUTIES_NUM,4+period*2)).options(transpose=True).value=df_out.iloc[:,period].values


    if x==-1: 
        print("\n\n!!!!!!!!!!!\nWARNING: NO SOLUTION COULD BE FOUND FOR WEEK",week_to_optimize,". NO CHANGES MADE.")
        input("press any key to continue...")
    else:
        print("\nWeek ",week_to_optimize," Solved","\n",df_out)

    
    
    return problem

if __name__ == "__main__":
    
    tk.Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
    filename = tk.filedialog.askopenfilename(initialdir=os.getcwd(), title="Select CC Roster file",
                                           filetypes=[("Excel Files", "*.xlsx")]) # show an "Open" dialog box and return the path to the selected file
    print(filename)
      
    try:
        wb = xw.Book(os.path.basename(filename))
        sheet = wb.sheets['Optimize']
        week_to_optimize=sheet.range(5,39).value
    except:
        print("Check error: Do you select a CC Roster file, with a sheet named 'Optimize' (O in caps) with a settings table in col AL, in the same folder as this program?")
        exit()
    
    if week_to_optimize=="All":
        for x in range(5): 
            print("Solving Week ",int(x),"\n","---------------------------------","\n\n")
            problem = model_problem(x+1)

    else:
        print("Solving Week ",int(week_to_optimize),"\n","---------------------------------","\n\n")
        problem=model_problem(int(week_to_optimize))
    