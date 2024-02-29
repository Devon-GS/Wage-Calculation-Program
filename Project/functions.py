import os
import importlib
from os import path
import database as db

import att_roster_times as ar
import att_clock_times as ac
import att_cal_hours as ath

import cas_roster_times as cr
import cas_clock_times as cc
import cas_cal_hours as cth

import excel_format as format

import payroll as pay


# importlib.reload(sys)

# ##########################################################
# ALL PROGRAM FUNCTIONS  
# ##########################################################

def program_init():
     db.db_init()

def wages_time_main_program():
        # REMOVE WAGES TIMES.XLSX
        wage_times = path.exists('Wage Times.xlsx')

        if wage_times == True:
            os.remove('Wage Times.xlsx')

        # CLEAN DATABASE
        db.clean_db()
        
        # ##########################################
        #               ATTENDENT
        # ##########################################

        # ATTENDENT ROSTER
        # Week One
        ar.db_init()
        ar.db_update_dates()
        ar.db_update_badges()
        ar.db_to_excel()

        # Week Two
        ar.db_atwo_init()
        ar.db_atwo_update_dates()
        ar.db_atwo_update_badges()
        ar.db_atwo_to_excel()

        # ATTENDENT CLOCK TIMES
        # Week One
        ac.recent_clock()
        ac.clock_times_collector()
        ac.att_clock_excel()

        # Week two
        ac.att_clock_excel_wt()

        # ATTENDENT TOTAL TIMES CALCULATION
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

        # ##########################################
        #               CASHIER
        # ##########################################

        # CASHIER ROSTER
        # Week One
        cr.db_cas_init()
        cr.db_update_cas_dates()
        cr.db_update_cas_badges()
        cr.db_to_excel()

        # Week Two
        cr.db_ctwo_init()
        cr.db_ctwo_update_dates()
        cr.db_ctwo_update_badges()
        cr.db_ctwo_to_excel()

        # CASHIER CLOCK TIMES
        # Week One
        cc.recent_clock()
        cc.clock_times_collector()
        cc.cashier_clock_excel()

        # Week two
        cc.cashier_clock_excel_wt()

        # CASHIER TOTAL TIMES CALCULATION
        # Week One
        cth.cas_times_weekone()
        cth.cas_public_weekone()
        cth.bak_cas_work()
        cth.cas_total_wo_hours()

        # Week Two
        cth.cas_times_weektwo()
        cth.cas_public_weektwo()
        cth.bak_cas_work_wt()
        cth.cas_total_wt_hours()

        # CASHIER TOTAL TIMES
        cth.cas_total_wo_db()
        cth.cas_total_wt_db()
        cth.cas_fortnight_total()

        # FORMAT EXCEL WOORKBOOK
        format.excel_format()

def recal_hours():
    # clean Totals table
    db.clean_db_recal()

    # Recalculate hours and push the db (Attendents)
    ath.att_times_weekone()
    ath.att_public_weekone()
    ath.att_total_wo_hours('yes')

    ath.att_times_weektwo()
    ath.att_public_weektwo()
    ath.att_total_wt_hours('yes')

    ath.att_total_wo_db()   
    ath.att_total_wt_db()
    ath.att_fortnight_total()

    # Recalculate hours and push the db (Cashiers)
    cth.cas_times_weekone()
    cth.cas_public_weekone()
    cth.bak_cas_work()
    cth.cas_total_wo_hours('yes')

    cth.cas_times_weektwo()
    cth.cas_public_weektwo()
    cth.bak_cas_work_wt()
    cth.cas_total_wt_hours('yes')

    cth.cas_total_wo_db()
    cth.cas_total_wt_db()
    cth.cas_fortnight_total()

def run_payroll():
     pay.payroll()
