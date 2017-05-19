# -*- coding: utf-8 -*-

import xlrd, xlwt, logging
from datetime import datetime

def ParseTxt(strFileNameConf):
    oFile = open(strFileNameConf, "r")
    dictConfConst = {}
    dictConfParam = {}

    #### transform txt into dict
    for strLine in oFile:
        strLine = strLine.strip()
        if len(strLine) == 0:
            oMyLog.debug('blank line is detected & ignored.')
        elif strLine[0] == '#':
            oMyLog.debug(strLine)
        else:
            strLine = strLine.replace('\n', '')
            listPair = strLine.split('=')

            if 'Constant' in listPair[0]:
                strKeyConst = listPair[0].strip()
                strValueConst = listPair[1].strip()
                dictLineConst = {strKeyConst: strValueConst}
                dictConfConst.update(dictLineConst)
                # print(dictConfConst)
            else:
                #### replace all the pre-defined constants
                for strKey in dictConfConst.keys():
                    if strKey in listPair[1]:
                        listPair[1] = listPair[1].replace(strKey, dictConfConst.get(strKey))
                strKeyParam =  listPair[0].strip()
                listValueParam = listPair[1].split(',')
                listTmp = []
                for strValueParam in listValueParam:
                    strValueParam = strValueParam.strip()
                    listTmp.append(strValueParam)
                listValueParam = listTmp
                dictLineParam = {strKeyParam: listValueParam}
                dictConfParam.update(dictLineParam)
                # print(dictConfParam)

    oFile.close()
    return dictConfParam, dictConfConst

def UpdateXls(dictConf):
    oMyLog.info('updating XLS...')
    nFiles = len(dictConfParam.get('FileNames'))
    oMyLog.info('No. of source files = ' + str(nFiles))
    oWorkbookTarget = xlwt.Workbook()
    oSheetTarget = oWorkbookTarget.add_sheet(strSheetNameTarget, cell_overwrite_ok=True)

    for nFile in range(nFiles):
        #### adding column headers by FileNames
        oSheetTarget.write(0, nFile + 1, dictConfParam.get('FileNames')[nFile])
        listTitles = ['Did Not Run', 'Pass', 'Fail', 'Automation']
        for nRow in range(len(listTitles)):
            #### adding row headers by listTitles
            oSheetTarget.write(nRow + 1, 0, listTitles[nRow])
            if (listTitles[nRow] != 'Automation'):
                UpdateXlsCells('TP', nFile, nRow, oSheetTarget)
            else:
                UpdateXlsCells('FEAT', nFile, nRow, oSheetTarget)

    oWorkbookTarget.save(strFileNameTargetXls)
    return oWorkbookTarget

def UpdateXlsCells(strFlag, nFile, nRow, oSheetTarget):
    oMyLog.info('updating XLS cells of Row_No.' + str(nRow) + ' from File_No.' + str(nFile) + '...')
    if strFlag == 'TP':
        strPaths, strFiles, strSheets, strAhchors = 'AbsolutePaths', 'FileNames', 'SheetNames', 'Anchors'
    elif strFlag == 'FEAT':
        strPaths, strFiles, strSheets, strAhchors = 'FeatPaths', 'FeatFileNames', 'FeatSheetNames', 'FeatAnchors'
    else:
        oMyLog.critical('invalid strFlag as: ' + strFlag)

    strFileLocation = dictConfParam.get(strPaths)[nFile] + dictConfParam.get(strFiles)[nFile]
    strSheetName = dictConfParam.get(strSheets)[nFile]
    listAnchor = dictConfParam.get(strAhchors)[nFile].split('|')
    strDelimiter = ', '
    listTmp = [strFileLocation, strSheetName, dictConfParam.get(strAhchors)[nFile]]
    oMyLog.debug('strFileLocation, strSheetName, listAnchor_NoSplit = ' + strDelimiter.join(listTmp))
    oWorkbookSource = xlrd.open_workbook(strFileLocation)
    oSheetSource = oWorkbookSource.sheet_by_name(strSheetName)

    #### calculating the cell location
    if strFlag == 'TP':
        nRowSource = int(listAnchor[0]) - 1 + nRow
    elif strFlag == 'FEAT':
        nRowSource = int(listAnchor[0]) - 1
    else:
        oMyLog.critical('invalid strFlag as: ' + strFlag)
    nColSource = int(listAnchor[1]) - 1
    nRowTarget = nRow + 1
    nColTarget = nFile + 1

    #### update oWorkbookTarget from oWorkbookSource
    oSheetTarget.write(nRowTarget, nColTarget, oSheetSource.cell(nRowSource, nColSource).value)
    strDelimiter = ', '
    listTmp = [str(nRowSource), str(nColSource), str(nRowTarget), str(nColTarget)]
    oMyLog.debug('nRowSource, nColSource, nRowTarget, nColTarget = ' + strDelimiter.join(listTmp))
    return strFileLocation

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
        logging.critical(strMsg)
    def error(self, strMsg):
        logging.error(strMsg)
    def warning(self, strMsg):
        logging.warning(strMsg)
    def info(self, strMsg):
        logging.info(strMsg)
    def debug(self, strMsg):
        logging.debug(strMsg)

if __name__ == '__main__':
    oMyLog = MyLog()
    # print(datetime.today())
    oMyLog.debug('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>')
    oMyLog.debug('Starting ParseConfig.py @ ' + str(datetime.today()))
    # strFilePath = 'D:\Technical_Document\44_Automation\Python\Scripts\Projects\CollectInfoToReport'
    strFileNameConf = 'Config.txt'
    strFileNameTargetXls = 'Report' + '_' + datetime.strftime(datetime.today(), '%Y%m%d_%H%M%S') + '.xls'
    strSheetNameTarget = 'Summary'
    # strFileNameTargetXls = 'Report.xls'
    dictConfParam, dictConfConst = ParseTxt(strFileNameConf)
    print(dictConfParam)
    print(dictConfConst)
    UpdateXls(dictConfParam)
    # print(datetime.today())
    oMyLog.debug('Ending ParseConfig.py @ '+ str(datetime.today()))
    oMyLog.debug('<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<\n')
