import gurobipy as gp
import numpy as np
import math
import pandas as pd 
from gurobipy import GRB

df1 = pd.read_excel('ExamRequirements_0326.xlsx') # Exam Information
df2 = pd.read_excel('StudentExam_0326.xlsx') # Student Information

Course_Duration = dict(zip(df1['Course Code'], df1['Duration']))
for course in Course_Duration:
    Course_Duration[course] = int(Course_Duration[course].strip('Minutes'))
Courses, Duration = gp.multidict(Course_Duration)

print("考试时长：")
print(Course_Duration)

ID_Course = df2.groupby('SID')['Course Code'].apply(set).to_dict() 
Course_ID = df2.groupby('Course Code')['SID'].apply(set).to_dict()

print("ID_Course:")
print(ID_Course)

NumberofCourses = len(Courses)
Timeslots = range(24)# everyday has three time slots 8:30, 13:30, 18:30(totally 3*8 slots)
CourseCapacity = {i:len(Course_ID[i]) for i in Courses}
Sharing = dict()
for i in Courses:
    for j in Courses:
        if i == j:
            Sharing[i,j] = 0
        else:
            Sharing[i,j] = len(Course_ID[i] & Course_ID[j])
CoursePair, Sharing = gp.multidict(Sharing)


Small_Courses = [course for course in Courses if CourseCapacity[course]<=100]

count = 0 
for i in Sharing :
    if Sharing[i] > 0 :
        count += 1
print(count)

TriSharing = dict()
for i in Courses:
    for j in Courses:
        for k in Courses:
            if (i == j) or (i == k) or (j == k):
                TriSharing[i,j,k] = 0
            else:
                TriSharing[i,j,k] = len(Course_ID[i] & Course_ID[j]& Course_ID[k])
TriCourse, TriSharing = gp.multidict(TriSharing)

t_830 = [0,3,6,9,12,15,18,21]
t_1330 = [i+1 for i in t_830]
t_1830 = [i+2 for i in t_830]
d1 = [0,1,2]
d2 = [3,4,5]
d3 = [6,7,8]
d4 = [9,10,11]
d5 = [12,13,14]
d6 = [15,16,17]
d7 = [18,19,20]
d8 = [21,22,23]

exam = gp.Model()
# Here we set the time limit as 20 min or 30 min to lower the gap. 
# totally we have 42w constrains, which takes us a long time to solve it.
exam.params.timelimit=10*60


start_time = exam.addVars(Courses,Timeslots,vtype='B',name='start')
robust_variable = exam.addVar(vtype='C',name='robust')
# y = exam.addVars(Courses,lb=0,name='y')

small_course_robust = exam.addVar(vtype='C',name='small_robust')

def special_constraint(p_course, p_candidate, p_suggest):
    l_sum = gp.LinExpr()
    for k in p_candidate:
        l_sum += start_time[p_course,k]
    if p_suggest == True:
        # exam.addConstr(gp.quicksum(start_time[p_course,k]==1 for k in p_candidate))

        exam.addConstr(l_sum==1)
    else:
        # exam.addConstr(gp.quicksum(start_time[p_course,k]==0 for k in p_candidate))
        exam.addConstr(l_sum==0)
    exam.update()

# the sum of exams in the weekend 
tail_sum = gp.LinExpr()
for k in d1+d2+d8:
    for course in Courses:
        tail_sum += start_time[course,k]*CourseCapacity[course]
# make this sum small ASAP

small_course_sum = dict()
for k in Timeslots:
    small_course_sum[k] = gp.LinExpr()
    for course in Small_Courses:
            small_course_sum[k] += start_time[course,k]*CourseCapacity[course]

# No three successive exams 
for i in Courses:
    for j in Courses:
        for p in Courses:
            if TriSharing[i,j,p]>0:
                for k in t_830:
                    exam.addConstr(start_time[i,k]+start_time[j,k+1]+start_time[p,k+2]<=2)

exam.addConstrs(start_time.sum(i,'*')==1 for i in Courses)

# add constrain that no exam on weekend night
for i in Courses :
    exam.addConstr(start_time[i,2] + start_time[i,5] + start_time[i,20] + start_time[i,23]  == 0)

# no exams overlaping
for i in Courses:
    for j in Courses:
        if Sharing[i,j]>0:
            for k in Timeslots:
                exam.addConstr(start_time[i,k]+start_time[j,k]<=1)


# maximum of student number is 1400 in one day
exam.addConstrs(gp.quicksum(start_time[i,j]*CourseCapacity[i] for i in Courses)<=1400 for j in Timeslots)
# exam.addConstrs(gp.quicksum(start_time[i,j]*CourseCapacity[i] for i in Courses) >=500 for j in [0,1])
# maximum of student number is 500 in the night
exam.addConstrs(gp.quicksum(start_time[i,3*j+2]*CourseCapacity[i] for i in Courses)<=500 for j in range(8))

exam.addConstrs(small_course_sum[k] <= small_course_robust for k in Timeslots)


# for course in Courses:
#     if CourseCapacity[course]>=150:
#         special_constraint(course,d1+d8,False)


# some special constrains     
special_constraint("FIN2010", [1], True) # 11TH MAY 13:30
special_constraint("DDA4210", [8], True) # 13TH MAY 18:30

# d1 = [0,1,2]FIN2010 1
# d2 = [3,4,5]
# d3 = [6,7,8]  DDA4210 8
# d4 = [9,10,11]  DMS2030 10
# d5 = [12,13,14]   STA3005 17
# d6 = [15,16,17] 
# d7 = [18,19,20] CSC3100 19
# d8 = [21,22,23]

# Course requirement
special_constraint("PHY1001", d7+d8, False)# Before 17th
special_constraint("FIN2010", d1+d2, True) # First two days
special_constraint("MAT2040", d5+d6, True) # in 15th or 16th
special_constraint('ECO3121', d1+d2+d3, True) #ASAP
special_constraint('BIO1001', [6], True) # 13th morning
special_constraint('MED2002', [0], True) # 11TH MAY 8:30-11:30 
special_constraint('MED2012', [3], True) # 12TH MAY 8:30-11:30
special_constraint('MED2040', [9], True) # 14TH MAY 8:30-10:30
special_constraint('MED2060', [6], True) # 13TH MAY 8:30-10:30
special_constraint('BIO2004', [12], True) # 15 MAY, IN THE MORNING
special_constraint('MED2030', [12], True) # 15TH MAY 8:30-10:30
special_constraint('MED2100', [13], True) # 15TH MAY 14:00-15:00
special_constraint('MED1002', [3], True) # 12TH MAY 9:00-12:00 
special_constraint('MED1012', [0], True) # 11TH MAY 9:00-12:00
special_constraint('MED1022', [6], True) # 13TH MAY 9:00-11:00
special_constraint('MED1032', [9], True)# 14TH MAY 9:00-11:00
special_constraint('PHM2002', [12],True) # 14th 8:30-10:30

# Scheduling requirement
special_constraint('MAT1002',[0,1,3,4]+d3,True)#ASAP 
special_constraint('PHY1001',[0,1,3,4]+d3,True)#ASAP 
special_constraint('FIN2010',[0,1,3,4]+d3,True)#ASAP
special_constraint('DMS2030',[0,1,3,4]+d3,True)#ASAP  
special_constraint('ECO2121',[0,1,3,4]+d3,True)#ASAP 




# exam.addConstr(evening_sum==0)
# exam.addConstrs(slot_sum[k]<=800 for k in Timeslots)
# exam.addConstrs(start_time.sum('*',j)>= robust_variable for j in Timeslots)
# exam.setObjective(10*robust_variable-1/8*tail_sum,sense=GRB.MAXIMIZE)
exam.setObjective(1/2*tail_sum+small_course_robust,sense=GRB.MINIMIZE)
exam.optimize()

def busy(schedule):
    if len(schedule)==1:
        return False
    elif len(schedule)==2:
        return schedule[1]-schedule[0]<=1
    else:
        for i in range(1,len(schedule)-1):
            if schedule[i]-schedule[i-1]==1 and schedule[i+1]-schedule[i]==1:
                return True
            

