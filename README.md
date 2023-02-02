# RosterShiftScheduler
This is a Roster shift allocation model for small teams of medical professionals with multiple specialized duties 
The python program code reads availability data from the excel file, and solves for the most 'balanced' allocation of duties across each employee for each week, subject to constraints.
It then writes the allocations back to the file. If minor changes are required, the code can be rerun after modifying the availability and it will allocate while attempting to minimize changes from the existing allocations

availabilty is specified in rows 59-76 of roster.xlsx. Any value in the leave calendar will be seen as unavailable. if the value contains AM or PM it will recognize it as unavailable only for that shift

Settings are specifed in cells AM4-AM15 of roster.xlsx

This project was done for a team of medical professionals with a specific manual method of rostering and with an existing Excel template that they wanted to continue using. 
if you hate the format of the excel file or it doesn't work for you, feel free to  modify this as required for your purposes. 


Credit: forked from https://github.com/lbiedma/shift-scheduling


Number of shifts = 10 (5 day work week with AM and PM Shifts)
Number of Duties = 4  (D1 and D2 are considered the hardest duties and hence need to balanced best, D3 is less hard, and D4 is easy)
Fixed Duties: these are a single head where clinicians who have a specific duty on a specific day need to be accomodated (for example a fixed clinic day)
NUmber of Employees = 12 (but can go up to 15 with the current excel file)

Hard Constraints
1. Every duty must be assigned to one and only one employee for each shift 
2. Employees cannot pull double duties on a single shift, except if explicitly allowed in the settings (in the excel file) and only can pull D4 (since it is easy)
3. Max number of total duties in a week (D1+D2+D3+D4) specified in settings
4. Max number of hard duties (D1+D2+D3) in a week specified in settings
5. Max number of the hardest duties (D1+D2) in a week specifed in settings
6. For those with a fixed duty on a particular day, do not assign D1+D2 on the same day (unless allowing in settings)

Soft Constraints: These apply penalties to the objective function to be minimized. ie. they are minimized as part of the problem 
7. Assign the same duty to the same employee in both AM and PM shifts (acc to weightage in settings)
8. if there are prior assignments in that week, make minimal changes to prior assignments (not currently a setting)

Minimize:
1. Difference between number of the hardest weekly duties assigned to each employee (D1+D2) and the average assigned for all employees
2. Difference between number of the hard weekly duties assigned to each employee (D3) and the average assigned for all employees
3. Difference between number of the easy weekly duties assigned to each employee (D4) and the average assigned for all employees

