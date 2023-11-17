from os import path
import database as db
import att_roster_times as awo
import att_clock_times as cwo
import att_cal_hours as atwo

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

# ATTENDENT WEEKONE ROSTER
awo.db_init()
awo.db_update_dates()
awo.db_update_badges()
awo.db_to_excel()
print('Attendent Weekone Finnished')

# ATTENDENT WEEKTWO ROSTER
awo.db_atwo_init()
awo.db_atwo_update_dates()
awo.db_atwo_update_badges()
awo.db_atwo_to_excel()
print('Attendent Weektwo Finnished')

# ATTENDENT WEEKONE CLOCK TIMES
cwo.recent_clock()
cwo.clock_times_collector()
cwo.att_clock_excel()
print('Attendent Weekone Clock Times Finnished')

# ATTENDENT WEEKTWO CLOCK TIMES

# ATTENDENT WEEKONE TIMES CALCULATION
atwo.att_times_weekone()
atwo.att_public_weekone()
atwo.att_total_wo_hours()
print('Attendent Weekone Times Finnished')

print('Wage Times.xlsx had printed and is ready for viewing')

