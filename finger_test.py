# -*- coding: utf-8 -*-
#
# tool   : test for finger log parser
# author : ralakus
# date   : 26.11.2013
#

import unittest
import finger
from datetime import datetime

# Used for returning some datetime object for time difference
EPOCH_TIME      = datetime(1970, 1, 1)
EPOCH_OVERWORK  = datetime(1970, 1, 1, 11, 0)
EPOCH_UNDERWORK = datetime(1970, 1, 1, 10, 0)


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
    
class TestSequenceFunctions(unittest.TestCase):

    def __user_log(self, finger_log, user_name):
        user_log = []
        try:
            for log in finger_log:
                try:
                    if log[DAILY_USER_NAME]== user_name:
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
                if log[DAILY_LOG_DATE].date() == check_date.date():
                    user_log.append(log)
            except IndexError:
                pass 
        return user_log
    
    def __hours_worked(self, daily_report, user, check_date): 
        log_tmp = self.__date_log(self.__user_log(daily_report, user), check_date)
        return log_tmp[0][DAILY_WORKED_HR]
       
    def __check_daily_status(self, worked_hours):
        if worked_hours > EPOCH_OVERWORK:
            return "Fazla"
        elif worked_hours == EPOCH_TIME:
            return "Gelmedi"
        elif worked_hours < EPOCH_UNDERWORK:
            return "Eksik"
        else:
            return "Normal" 


    def setUp(self):
        pass

#    def test_sabah_giris_ayni_gun_aksam_cikis(self):
#        log_out = []
#        loader = finger.ExcelLoader()
#        loader.generate_list(loader.open_book("test_sabah_giris_ayni_gun_aksam_cikis.xls"), log_out)
#        loader.generate_list(loader.open_book("test_sabah_giris_ayni_gun_aksam_cikis.xls"), log_out)
#        user_search = finger.Finger(log_out)
#        hours_worked =  user_search.hours_worked(u"HER GÜN",'2013-11-08')
#        actual_worked = datetime(2013, 11, 8, 8, 28, 50)     
#        self.assertTrue(str(hours_worked.time()) == str(actual_worked.time()))  
     
#    def test_gelinmeyen_gun(self):
#        log_out = []
#        loader = finger.ExcelLoader()
#        loader.generate_list(loader.open_book("test_eksik_saat_calisma.xls"), log_out)
#        loader.generate_list(loader.open_book("test_eksik_saat_calisma.xls"), log_out)
#        user_search = finger.Finger(log_out)
#        worked_hours =  user_search.hours_worked(u"HER GÜN",'2013-11-08')
#        self.assertTrue(user_search.check_daily_status(worked_hours) == "Gelmedi")  

    def test_eksik_saat_calisma(self):
        log_out = []
        loader = finger.ExcelLoader()
        loader.generate_list(loader.open_book("test_eksik_saat_calisma.xls"), log_out)
        user_search = finger.Finger(log_out)
        daily_report = user_search.get_daily_report()
        date_test = datetime(2013, 10, 26)
        worked_hours =  self.__hours_worked(daily_report, u"HER GÜN", date_test)
        self.assertTrue(self.__check_daily_status(worked_hours) == 'Eksik') 
    
     
if __name__ == '__main__':
    unittest.main()