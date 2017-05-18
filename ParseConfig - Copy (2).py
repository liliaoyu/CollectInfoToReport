# -*- coding: utf-8 -*-

import os, re, xlrd, xlwt, logging
from datetime import datetime

def ParseTxt(strFileNameConf):
    oFile = open(strFileNameConf, "r")
    dictConf = {}
    strAbsolutePathDefault = ''
    #### transform txt into dict
    for strLine in oFile:
        strLine = strLine.replace('\n', '')
        listPair = strLine.split('=')
        if listPair[0] == 'AbsolutePathDefault':
            strAbsolutePathDefault = listPair[1]
            oMyLog.debug('strAbsolutePathDefault = ' + strAbsolutePathDefault)
        if 'AbsolutePath' in listPair[1]:
            listPair[1] = listPair[1].replace('AbsolutePath', strAbsolutePathDefault)
        listPair[1] = listPair[1].split(',')
        dictLine = {listPair[0]: listPair[1]}
        dictConf.update(dictLine)
        # print(dictConf)
    oFile.close()
    return dictConf

def UpdateXls(dictConf):
    nFiles = len(dictConf.get('FileNames'))
    oMyLog.info('No. of source files = ' + str(nFiles))
    oWorkbookTarget = xlwt.Workbook()
    oSheetTarget = oWorkbookTarget.add_sheet(strSheetNameTarget, cell_overwrite_ok=True)

    for nIndex in range(nFiles):
        strFileLocation = dictConf.get('AbsolutePaths')[nIndex] + dictConf.get('FileNames')[nIndex]
        strSheetName = dictConf.get('SheetNames')[nIndex]
        listAnchor = dictConf.get('Anchors')[nIndex].split('|')
        oMyLog.debug('strFileLocation, strSheetName, listAnchor_NoSplit = ' \
                    + strFileLocation + ', ' + strSheetName + ', ' + dictConf.get('Anchors')[nIndex])

        #### update oWorkbookTarget from oWorkbookSource
        oWorkbookSource = xlrd.open_workbook(strFileLocation)
        oSheetSource = oWorkbookSource.sheet_by_name(strSheetName)
        oSheetTarget.write(0, nIndex + 1, dictConf.get('FileNames')[nIndex])
        strItems = ['Did Not Run', 'Pass', 'Fail']
        for nRow in range(3):
            nRowSource = int(listAnchor[0]) + nRow
            nColSource = int(listAnchor[1])
            nRowTarget = nRow + 1
            nColTarget = nIndex + 1
            oMyLog.debug('nRowSource, nColSource, nRowTarget, nColTarget = ' \
                        + str(nRowSource) + ', ' + str(nColSource) + str(nRowTarget) + ', ' + str(nColTarget))
            oSheetTarget.write(nRowTarget, 0, strItems[nRow])
            oSheetTarget.write(nRowTarget, nColTarget, oSheetSource.cell(nRowSource, nColSource).value)

    oWorkbookTarget.save(strFileNameTargetXls)
    return oWorkbookTarget

class MyLog(object):
    def __init__(self):
        logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                        datefmt='%a, %d %b %Y %H:%M:%S',
                        filename='PyLog.log',
                        filemode='a+')
        #### adding output on the screen
        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
        console.setFormatter(formatter)
        logging.getLogger('').addHandler(console)

    def critical(self, strMsg):
        logging.critical(strMsg + '\n')
    def error(self, strMsg):
        logging.error(strMsg + '\n')
    def warning(self, strMsg):
        logging.warning(strMsg + '\n')
    def info(self, strMsg):
        logging.info(strMsg + '\n')
    def debug(self, strMsg):
        logging.debug(strMsg + '\n')

if __name__ == '__main__':
    oMyLog = MyLog()
    # print(datetime.today())
    oMyLog.debug('>>>>>>>> Starting ParseConfig.py @ ' + str(datetime.today()))
    # strFilePath = 'D:\Technical_Document\44_Automation\Python\Scripts\Projects\CollectInfoToReport'
    strFileNameConf = 'Config.txt'
    strFileNameTargetXls = 'Report' + '_' + datetime.strftime(datetime.today(), '%Y%m%d_%H%M%S') + '.xls'
    strSheetNameTarget = 'Summary'
    # strFileNameTargetXls = 'Report.xls'
    dictConf = ParseTxt(strFileNameConf)
    print(dictConf)
    UpdateXls(dictConf)
    # print(datetime.today())
    oMyLog.debug('<<<<<<<< Ending ParseConfig.py @ '+ str(datetime.today()))
