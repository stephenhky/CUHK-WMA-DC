# -*- coding: utf-8 -*-
"""
Created on Sun Jun  9 15:43:44 2013

@author: hok1
"""

from google_spreadsheet.api import SpreadsheetAPI
import getpass
from gdata.service import BadAuthentication

DatabaseCUHKWMAID = '0Aq8P97Yci1OCdE1UWU5uZDN5MG1ucE1yVklWUWJLM2c'

class GoogleSpreadsheet:
    @staticmethod
    def getLoginInfo():
        user_gmail = raw_input('You Gmail address? ')
        user_password = getpass.getpass('Your Gmail password? ')
        return user_gmail, user_password
        
    @staticmethod
    def getWorksheetsDict(worksheets):
        worksheetDict = {}
        for name, code in worksheets:
            worksheetDict[name] = code
        return worksheetDict

    @staticmethod        
    def wrapRowsData(rows):
        headerRow = rows[0]
        headerDict = {}
        for rowCode in headerRow:
            headerDict[headerRow[rowCode]] = rowCode
            
        data = []
        for row in rows[1:]:
            item = {}
            for name in headerDict:
                item[name] = row[headerDict[name]]
            data.append(item)    
        return data
        
class DatabaseCUHKWMA:
    def __init__(self, username, password):
        self.api = SpreadsheetAPI(username, password, '')
        worksheets = self.api.list_worksheets(DatabaseCUHKWMAID)
        self.sheetsDict = GoogleSpreadsheet.getWorksheetsDict(worksheets)
        sheet = self.api.get_worksheet(DatabaseCUHKWMAID,
                                       self.sheetsDict['Current Members'])
        #self.rows = sheet.get_rows()
        #self.data = GoogleSpreadsheet.wrapRowsData(self.rows)
        self.data = sheet.get_rows()
        
if __name__ == '__main__':
    try:
        username, password = GoogleSpreadsheet.getLoginInfo()
        outputfilename = raw_input('Output file name=? ')
        outf = open(outputfilename, 'wb')
        dbCUHK = DatabaseCUHKWMA(username, password)
        for item in sorted(dbCUHK.data, key=lambda item: item['lastname']):
            lastname = item['lastname'] if not (item['lastname'] is None) else ''
            firstname = item['firstname'] if not (item['firstname'] is None) else ''
            chinesename = item['chinesename'] if not (item['chinesename'] is None) else ''
            email = item['emailaddress'] if not (item['emailaddress'] is None) else ''
            
            lastname = lastname.strip()
            firstname = firstname.strip()
            chinesename = chinesename.strip()
            email = email.strip()
            rowstr = unicode(firstname+' '+lastname+' '+chinesename+' <'+email.strip()+'>,\n')
            
            outf.write(rowstr.encode('utf8'))
        outf.close()
    except BadAuthentication:
        print 'Invalid e-mail or password.'
