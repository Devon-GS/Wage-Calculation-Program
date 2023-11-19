from os import path
import database as db
import att_roster_times as ar
import att_clock_times as ac
import att_cal_hours as ath

import excel_format as format

# database_file = path.exists('wageTimes.db')

# user_input = ''

# while True:

#     print('Select one of the following to proceed:')
#     print('Type 1: Initiate Database')
#     print('Type 2: New Fortnight Hour Calculation')
#     print()

#     user_input = input('Selection = ')

#     if user_input == '1':
#         if database_file:
#             user = input('Database already exists...type y to continue or n to exit: ')
#             if user == 'y':
#                 print('yes database is init')
#             elif user == 'n':
#                 break
#             else:
#                 print('database created')
#                 continue
#     elif user_input == '2':
#         print('##########################################################################')
#         print('Have you edited the Attend and Cashier roster to current fortnight?')
#         print('Are all public holidays up todate?')
#         print('Are all badges up todate?')
#         print()
#         user = input('If you are ready to proceed type y or n if you are not ready: ')
#         if user == 'y':
#             print('ready')
#             break
#         break
        


# CLEAN DATABASE
db.clean_db()
print('Database Cleaned')

# ##########################################
#               ATTENDENT
# ##########################################

# ATTENDENT ROSTER
# Week One
ar.db_init()
ar.db_update_dates()
ar.db_update_badges()
ar.db_to_excel()
print('Attendent Weekone Finnished')

# Week Two
ar.db_atwo_init()
ar.db_atwo_update_dates()
ar.db_atwo_update_badges()
ar.db_atwo_to_excel()

print('Attendent Weektwo Finnished')

# ATTENDENT CLOCK TIMES
# Week One
ac.recent_clock()
ac.clock_times_collector()
ac.att_clock_excel()

# Week two
ac.att_clock_excel_wt()

print('Attendent Weekone Clock Times Finnished')


# ATTENDENT TIMES CALCULATION
# Week One
ath.att_times_weekone()
ath.att_public_weekone()
ath.att_total_wo_hours()

# Week Two
ath.att_times_weektwo()
ath.att_public_weektwo()
ath.att_total_wt_hours()

# ATTENDENT TOTAL TIMES
ath.att_total_wo_db()   
ath.att_total_wt_db()
ath.att_fortnight_total()

print('Attendent Weekone Times Finnished')

# FORMAT EXCEL WOORKBOOK
format.excel_format()

print('Wage Times.xlsx had printed and is ready for viewing')

