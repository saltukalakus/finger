# -*- coding: utf-8 -*-
#
# tool   : finger log parser for entry of personnel
# author : ralakus
# date   : 26.11.2013
#

from xlrd import open_workbook, xldate
from datetime import datetime, timedelta
import re

# Constants based on excel data structure
user_entry = u"ANA GİRİŞ"
user_exit  = u"ANA ÇIKIŞ"
LOG_USER_NAME      = 0
LOG_USER_SURNAME   = 1
LOG_GATE_ACTION    = 2
LOG_DATE           = 3

# Gate enter/exit actions
MAIN_GATE_EXIT      = 0
MAIN_GATE_ENTRY     = 1
OTHER_GATES         = 2

# Used for returning some datetime object for time difference
EPOCH_TIME      = datetime(1970, 1, 1)
EPOCH_OVERWORK  = datetime(1970, 1, 1, 11, 0)
EPOCH_UNDERWORK = datetime(1970, 1, 1, 10, 0)


# Helper class to generate the list form excel while taking care of data structure
#
# Output format:
#    Name
#    Surname
#    Main gate entry/exit (1,0) rest of gates (2)
#    Date and time
class ExcelLoader(object):

    # Open the excel file
    def open_book(self, file_name):
        return open_workbook(file_name, on_demand=True)
        
    # Generate the list
    def generate_list(self, book, log_out):
        sheet = book.sheet_by_index(0)
        rowIndex = -1
    
        try:
            for cell in sheet.col(0):
                rowIndex += 1
                cells = sheet.row(rowIndex)
                row_tmp = []
                colIndex = -1
                
                # iterate on the row
                for cell in cells:
                    if rowIndex > 1: # skip first 2 row 
                        colIndex +=1
                        
                        # user name/surname
                        if (colIndex < LOG_GATE_ACTION):                
                            row_tmp.append(cell.value)
                        # entry exit of main gate
                        elif colIndex == LOG_GATE_ACTION:
                            if cell.value == user_entry:
                                row_tmp.append(MAIN_GATE_ENTRY)
                            elif cell.value == user_exit:
                                row_tmp.append(MAIN_GATE_EXIT)
                            else:
                                row_tmp.append(OTHER_GATES)
                        # date - time log
                        elif colIndex == LOG_DATE:
                            try:
                                datetime_value = datetime(*xldate.xldate_as_tuple(cell.value, 0))
                                row_tmp.append(datetime_value)
                            except ValueError:
                                pass
                        
                log_out.append(row_tmp)
                
        except TypeError:
            pass
    
        book.unload_sheet(0)        

# Filters useful finger logs, finds unique users, finds date range in the log
class FingerFilter(object):

    def __init__(self, finger_log):
        self.m_filtered_log = self.__filter_other_gates_log(finger_log)
        self.m_user_names = self.__get_name_list(self.m_filtered_log)
        self.m_first_day, self.m_last_day = self.__get_day_range(self.m_filtered_log)
        self.m_first_week, self.m_last_week = self.__get_week_range(self.m_first_day, self.m_last_day)
        self.m_filtered_log = self.__filter_to_single_entry_and_exit(self.m_filtered_log)

    def __contains_digits(self, d):
        try:
            return bool(re.compile('\d').search(d))
        except TypeError:
            pass
        return True
                
    # Extract unique user names from log
    def __get_name_list(self, finger_log):
        name_list = []
        for log in finger_log:
            name_found = False
            for name in name_list:
                if name == log[LOG_USER_NAME] or self.__contains_digits(log[LOG_USER_NAME]) == True:
                    name_found = True
                    break
            if name_found == False:
                try:
                    name_list.append(log[LOG_USER_NAME])
                    #print "New user: " + log[LOG_USER_NAME]
                except IndexError:
                    pass
                    
        return name_list
    
    # Extract the day range from log
    def __get_day_range(self, finger_log):
        first_day = datetime(2099, 01, 01)
        last_day  = datetime(2013, 01, 01)
        for log in finger_log:
            if log[LOG_DATE].date() > last_day.date():
                last_day = log[LOG_DATE]
            if log[LOG_DATE].date() < first_day.date():
                first_day = log[LOG_DATE]                

        #print "Days begin: " + str(first_day) + " end date: " + str(last_day)
        return first_day, last_day
    
    # Get week range from day range
    def __get_week_range(self, first_day, last_day):
        first_week = datetime.date(first_day).isocalendar()[1]
        last_week = datetime.date(last_day).isocalendar()[1]
        return first_week, last_week 
        
        
    # Filter gate entries other then the main gate
    def __filter_other_gates_log(self, finger_log):
        log_out = []
        for log in finger_log:
            try:
                if log[LOG_GATE_ACTION] != OTHER_GATES:
                    log_out.append(log)
            except IndexError:
                pass
        return log_out
                
    # Filter daily log to single entry and exit
    def __filter_to_single_entry_and_exit(self, finger_log):
        log_out = []
        day = self.m_first_day
        
        while day.date() <= self.m_last_day.date():

            for user in self.m_user_names:
                earliest_entry     = datetime(2099,1,1,23,59)
                store_entry_log    = False    
                earliest_entry_log = 0

                latest_exit        = datetime(2013,1,1,0,0)
                store_exit_log     = False
                latest_exit_log    = 0
                 
                for log in finger_log:              
                    if log[LOG_USER_NAME] == user and log[LOG_DATE].date() == day.date():
                        # find the earliest entry
                        if log[LOG_DATE].time() < earliest_entry.time() and log[LOG_GATE_ACTION] == MAIN_GATE_ENTRY:
                            earliest_entry = log[LOG_DATE]
                            earliest_entry_log = log
                            store_entry_log = True
                                  
                        # find the latest exit
                        if log[LOG_DATE].time() > latest_exit.time() and log[LOG_GATE_ACTION] == MAIN_GATE_EXIT:
                            latest_exit = log[LOG_DATE]
                            latest_exit_log = log
                            store_exit_log = True
                 
                if store_entry_log == True:
                    store_entry_log = False
                    log_out.append(earliest_entry_log)
          
                if store_exit_log == True:
                    store_exit_log = False
                    log_out.append(latest_exit_log)           
                               
            next_day = timedelta(days=1)                    
            day += next_day
            
        return log_out
                
    def get_log(self):
        return self.m_filtered_log
    
    def get_user_names(self):
        return self.m_user_names
    
    def get_first_date(self):
        return self.m_first_day

    def get_last_date(self):
        return self.m_last_day
    
    def get_first_week(self):
        return self.m_first_week
    
    def get_last_week(self):
        return self.m_last_week
    
class Finger(object):

    # Name: Date: Week Number: Worked Hours: Status:  
    DAILY_USER_NAME  =  0
    DAILY_LOG_DATE   =  1
    DAILY_WEEK_NB    =  2
    DAILY_WORKED_HR  =  3
    DAILY_STATUS     =  4

    # Name: Week number: Total hours worked:
    WEEKLY_USER_NAME =  0
    WEEKLY_WEEK_NB   =  1
    WEEKLY_TOTAL_HR  =  2
    
    def __init__(self, finger_log):
        self.m_filter        = FingerFilter(finger_log)
        self.m_log           = self.m_filter.get_log()
        self.m_user_names    = self.m_filter.get_user_names()
        self.m_first_day     = self.m_filter.get_first_date()
        self.m_last_day      = self.m_filter.get_last_date()
        self.m_first_week    = self.m_filter.get_first_week()
        self.m_last_week     = self.m_filter.get_last_week()
        self.m_daily_report  = self.__generate_daily_report()
        self.m_weekly_report = self.__generate_weekly_report()

    def __user_log(self, finger_log, user_name):
        user_log = []
        try:
            for log in finger_log:
                try:
                    if log[LOG_USER_NAME]== user_name:
                        user_log.append(log)
                except IndexError:
                    pass 
        except TypeError:
            pass
        return user_log
    
    def __date_log(self, finger_log, check_date):
        user_log = []
        for log in finger_log:
            try:
                if log[LOG_DATE].date() == check_date.date():
                    user_log.append(log)
            except IndexError:
                pass 
        return user_log
    
    def __hours_worked(self, user, check_date): 
        log_tmp = self.__date_log(self.__user_log(self.m_log, user), check_date)
        date_out = date_in = EPOCH_TIME

        for log in log_tmp:
            #print log
            if log[LOG_GATE_ACTION] == MAIN_GATE_EXIT:
                date_out = log[LOG_DATE]
            elif log[LOG_GATE_ACTION] == MAIN_GATE_ENTRY:
                date_in = log[LOG_DATE]
        
        if date_out == EPOCH_TIME or date_in == EPOCH_TIME:
            print "User missed to log in or out using finger" 
            worked_hours = EPOCH_TIME
        else:  
            new_diff =  date_out - date_in
            worked_hours = EPOCH_TIME + new_diff
        return worked_hours
       
    def __check_daily_status(self, worked_hours):
        if worked_hours > EPOCH_OVERWORK:
            return "Fazla"
        elif worked_hours == EPOCH_TIME:
            return "Gelmedi"
        elif worked_hours < EPOCH_UNDERWORK:
            return "Eksik"
        else:
            return "Normal" 

    # Name: Date: Week Number: Worked Hours: Status:  
    def __generate_daily_report(self):
        work_hour_log_out = []
        day = self.m_first_day
        
        while day.date() <= self.m_last_day.date():   
            for user in self.m_user_names:
                row_tmp = []
                row_tmp.append(user)
                row_tmp.append(day)
                row_tmp.append(datetime.date(day).isocalendar()[1])
                worked_hours = self.__hours_worked(user, day)
                row_tmp.append(worked_hours)
                row_tmp.append(self.__check_daily_status(worked_hours))
                work_hour_log_out.append(row_tmp)
            next_day = timedelta(days=1)                    
            day += next_day
   
        return work_hour_log_out

    # Name: Week number: Total hours worked:
    def __generate_weekly_report(self):
        weekly_work_log = []
        week = self.m_first_week
        
        for user in self.m_user_names:
            while week <= self.m_last_week:
                total_worked = EPOCH_TIME
                for log in self.m_daily_report:
                    if user == log[self.DAILY_USER_NAME] and week == log[self.DAILY_WEEK_NB]:
                        total_worked += (log[self.DAILY_WORKED_HR] - EPOCH_TIME)
                
                if total_worked != EPOCH_TIME:
                    row_tmp = []
                    row_tmp.append(user)
                    row_tmp.append(week)
                    row_tmp.append(total_worked)
                    weekly_work_log.append(row_tmp)
                week += 1
        return weekly_work_log
    
    def get_weekly_report(self):
        return self.m_weekly_report
    
    def get_daily_report(self):
        return self.m_daily_report

class FingerLogger():
    def __init__(self, finger_log):
        self.m_finger  = Finger(finger_log)
        
    def generate_report(self, file_name):
        self.m_finger.get_daily_report()
        self.m_finger.get_weekly_report()
        
        
if __name__ == "__main__":
    log_out = []
    loader = ExcelLoader()
    loader.generate_list(loader.open_book('simple.xls'), log_out)
    finger_logger = FingerLogger(log_out)
    finger_logger.generate_report()
