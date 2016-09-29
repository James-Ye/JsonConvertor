from collections import OrderedDict
import xlrd
import json
import os
from os.path import join

#########################################################################
def getPurKey(key):
    index_str = 1
    while key.find("[") != -1:
        idx = key.find("[")
        idx2 = key.find("]")
        if idx != -1 and idx2 != -1:
            key1 = key[0:idx]
            index_str = key[idx+1:idx2]
            key2 = key[idx2+1:len(key)]
            key = key1 + key2
    return (key,int(index_str))
#########################################################################

def copyItem(objsrc):
    objdes = None
    if str(type(objsrc)) == "<class 'collections.OrderedDict'>":
        objdes = OrderedDict()
        items = list(objsrc.items())
        for i in range(0, len(objsrc)):
            objkey = items[i][0]
            objvalue = items[i][1]
            objdes.setdefault(objkey,copyItem(objvalue))
    elif str(type(objsrc)) == "<class 'dict'>":
        objdes = {}
        items = list(objsrc.items())
        for i in range(0, len(objsrc)):
            objkey = items[i][0]
            objvalue = items[i][1]
            objdes.setdefault(objkey,copyItem(objvalue))
    elif str(type(objsrc)) == "<class 'list'>":
        objdes = []
        for i in range(0, len(objsrc)):
            objdes.append(copyItem(objsrc[i]))
    else:
        objdes = objsrc
    return objdes
        

###########################################################################################
def delete_file_folder(src):
    if os.path.isfile(src):
        try:
            os.remove(src)
        except:
            pass
 
    elif os.path.isdir(src):
        for item in os.listdir(src):
            itemsrc=os.path.join(src,item)
            delete_file_folder(itemsrc) 
        try:
            os.rmdir(src)
        except:
            pass

###########################################################################################
def getDictionary(fileName, dict_element, list_list):
    bk = xlrd.open_workbook(fileName)
    shxrange = range(bk.nsheets)
    sh = bk.sheet_by_name("dictionary")
    #获取行数
    nrows = sh.nrows
    #获取列数
    ncols = sh.ncols

    statuscol = 1
    for j in range(1,ncols):
        cell_value = sh.cell_value(0,j)
        if cell_value == "Status":
            statuscol = j
            break

    row_list = []
    last_value = ""

    for i in range(1,nrows):  
        for j in range(0,ncols):
            cell_value = sh.cell_value(i,j)
            if str(type(cell_value)) == "<class 'str'>":
                if j == statuscol:
                    if cell_value == "◆":
                        key = sh.cell_value(i,j+1)
                        if key != '':
                            value = ""
                            for strRow in row_list:
                                value += strRow
                                value += "."
                            value += last_value
                            list_list.append((key, value))
                    elif cell_value == "●":
                        key = sh.cell_value(i,j+1)
                        if key != '':
                            value = ""
                            for strRow in row_list:
                                value += strRow
                                value += "."
                            value += last_value
                            dict_element.setdefault(key, value)
                    break
                elif j < statuscol:
                    if cell_value != "":
                        level = len(row_list)
                        if j == level:
                            last_value = cell_value
                        elif j > level:
                            row_list.append(last_value)
                            last_value = cell_value
                        elif j < level:
                            for n in range(j,level):
                                del row_list[-1]
                            last_value = cell_value
    return
###########################################################################################

######################################################################################
def delvoiddict(dictionary):
    count = len(dictionary)
    while count > 0:
        list_keys = list(dictionary.keys())
        curKey = list_keys[count -1]
        value = dictionary[curKey]
        if str(type(value)) == "<class 'collections.OrderedDict'>":
            delvoiddict(value)
            if len(value) == 0:
                del dictionary[curKey]
        elif str(type(value)) == "<class 'list'>":
            list_child = list(value)
            list_count = len(list_child)
            while list_count > 0:
                list_value = list_child[list_count -1]
                if str(type(list_value)) == "<class 'collections.OrderedDict'>":
                    delvoiddict(list_value)
                    if len(list_value) == 0:
                        del list_child[list_count -1]
                    list_count -= 1
                else:
                    continue

            if len(value) == 0:
                del dictionary[curKey]

        
        count -= 1
    return
######################################################################################

######################################################################################
def readTemplate(fileName):
    bk = xlrd.open_workbook(fileName)
    shxrange = range(bk.nsheets)
    sh = bk.sheet_by_name("dictionary")
    #获取行数
    nrows = sh.nrows
    #获取列数
    ncols = sh.ncols
    #获取第一行第一列数据 
    cell_value = sh.cell_value(1,1)

    statuscol = 1
    for j in range(1,ncols):
        cell_value = sh.cell_value(0,j)
        if cell_value == "Status":
            statuscol = j
            break

    dict_main = OrderedDict()
    cell_value = ""
    curlevel = 0

    preDict = dict_main
    preKey = ''
    savedKey = ''
    isValue = True

    current_obj_list = []

    isDictBegin = False

    for i in range(1,nrows):  
        for j in range(0,ncols):
            cell_value = sh.cell_value(i,j)
            if str(type(cell_value)) != "<class 'str'>":
                continue
            if j == statuscol:
                if cell_value == "◆":
                    d = OrderedDict()
                    d.setdefault('XXX','')
                    preDict[preKey] = [d]
                    preDict = d
                    preKey = 'XXX'
                    current_obj_list.append(preDict)
                elif cell_value == "○":
                    del preDict[preKey]
                    preKey = savedKey
                    savedKey = ''
                break
            elif j < statuscol:
                if cell_value != '':
                    if j == 0:
                        preDict = dict_main
                        preDict.setdefault(cell_value,'')
                        preKey = cell_value
                        savedKey = ''
                        current_obj_list.clear()
                        current_obj_list.append(preDict)
                    elif j == curlevel:
                        preDict.setdefault(cell_value,'')
                        savedKey = preKey
                        preKey = cell_value
                    elif j > curlevel:
                        if preKey == 'XXX':
                            del preDict[preKey]
                            preDict.setdefault(cell_value,'')
                            preKey = cell_value
                        else:
                            d = OrderedDict()
                            d.setdefault(cell_value,'')
                            preDict[preKey] = d
                            preDict = d
                            savedKey = ''
                            preKey = cell_value
                            current_obj_list.append(preDict)
                    else:
                        list_count = len(current_obj_list)
                        if j < list_count:
                            preDict = current_obj_list[j]
                            for n in range(j+1,list_count):
                                del current_obj_list[-1]
                            preDict.setdefault(cell_value,'')
                            savedKey = ''
                            preKey = cell_value
                            
                    curlevel = j
            else:
                break

    delvoiddict(dict_main)

    return dict_main

######################################################################################


def getheadlist(list_list):
    headlist = []
    for i in range(0,len(list_list)):
        headlist.append(list_list[i][0])
    return headlist

#########################################################################
def processDictionary(dictionary, key, list_list, oldJsonList, json_list_list, list_index):
    RPHList = ["TransactionInput.PricingInput.Pnr.Segments",
    "TransactionInput.PricingInput.Pnr.Passenger"]
    headlist = getheadlist(list_list);
    count = len(dictionary)
    isInList = False
    for i in range(0, count):
        list_keys = list(dictionary.keys())
        itemKey = list_keys[i]
        value = dictionary[itemKey]
        if key == "":
            curKey = key + itemKey
        else:
            curKey = key + "." + itemKey

        if curKey in headlist:
            isInList = True
            
        list_count = 1
        if str(type(value)) == "<class 'dict'>":
            dict_child = dict(value)
            processDictionary(dict_child, curKey, list_list, oldJsonList, json_list_list, list_index)
        elif str(type(value)) == "<class 'list'>":
            list_child = list(value)
            list_count = len(list_child)
            for j in range(0, list_count):
                list_value = list_child[j]
                if str(type(list_value)) == "<class 'dict'>":
                    if isInList:
                        list_index = j + 1
                    dict_child = dict(list_value)
                    processDictionary(dict_child, curKey, list_list, oldJsonList, json_list_list, list_index)
                else:
                    if list_index > 1:
                        curKey = key + "[%d]"%(list_index) + "." + itemKey

                    if key in RPHList:
                        indexKey = ""
                        if list_index > 1:
                            indexKey = key + "[%d]"%(list_index) + "." + "Index"
                        elif list_index == 1:
                            indexKey = key + "." + "Index"
                        oldJsonList.append((indexKey,"%d"%(list_index)))

                    tup = (curKey, list_child)
                    oldJsonList.append(tup)
                    break
            ist_index = 0
        else:
            if list_index > 1:
                curKey = key + "[%d]"%(list_index) + "." + itemKey

            if key in RPHList:
                indexKey = ""
                if list_index > 1:
                    indexKey = key + "[%d]"%(list_index) + "." + "Index"
                elif list_index == 1:
                    indexKey = key + "." + "Index"
                oldJsonList.append((indexKey,"%d"%(list_index)))
                
            tup = (curKey, value)
            oldJsonList.append(tup)
            
        if curKey in headlist:
            idx = headlist.index(curKey)
            listtup =  (list_list[idx], list_count)
            json_list_list.append(listtup)
    return
#########################################################################
def getAgentContext(keyName):
    context = ""
    if keyName == "TransactionInput.PricingInput.Agent.IataNum":
        context = "IATANumber"
    elif keyName == "TransactionInput.PricingInput.Agent.CRS":
        context = "CRSCode"
    elif keyName == "TransactionInput.PricingInput.Agent.DeptCode":
        context = "DepartmentCode"
    return context

def getDiagnosticContext(keyName):
    context = ""
    if keyName == "TransactionInput.PricingInput.Options.Diagnostic.DiagnosticType.SliceAndDice":
        context = "SliceAndDice"
    elif keyName == "TransactionInput.PricingInput.Options.Diagnostic.DiagnosticType.Category":
        context = "RuleValidation"
    elif keyName == "TransactionInput.PricingInput.Options.Diagnostic.DiagnosticType.FareRetrieve":
        context = "FareRetrieval"
    elif keyName == "TransactionInput.PricingInput.Options.Diagnostic.DiagnosticType.YQYR":
        context = "YQYR"
    return context
#########################################################################
def ConverToTrueFalse(strYN):
    context = False
    if strYN == "Y":
        context = True
    return context
#########################################################################
def converYNToTrueFalse(oldJsonList):
    YN_list = ["TransactionInput.PricingInput.Agent.IsAgency",
    "TransactionInput.PricingInput.Pnr.Segments.IsForceStopover",
    "TransactionInput.PricingInput.Pnr.Segments.IsForceConnection",
    "TransactionInput.PricingInput.Options.AllEndosAppl",
    "TransactionInput.PricingInput.Options.IsEtkt",
    "TransactionInput.PricingInput.Options.FbcInHfcAppl",
    "TransactionInput.PricingInput.Options.FboxCurOverride",
    "TransactionInput.PricingInput.Options.InterlineOverride",
    "TransactionInput.PricingInput.Options.NetFareAppl",
    "TransactionInput.PricingInput.Options.IsBestbuy",
    "TransactionInput.PricingInput.Options.TaxDetailAppl",
    "TransactionInput.PricingInput.Options.FilterPtcAppl",
    "TransactionInput.PricingInput.Options.TaxSummaryAppl",
    "TransactionInput.PricingInput.Options.PrivateNegoFaresAppl",
    "TransactionInput.PricingInput.Pnr.Segments.IsOpen",
    "TransactionInput.PricingInput.Options.YqyrOnly",
    "TransactionInput.PricingInput.Options.TaxOnly"]
    useIsForceStopover = []
    modifyList = []
    removeList = []
    for i in range(0, len(oldJsonList)):
        tup = oldJsonList[i]
        tupkey = getPurKey(tup[0])
        oldkey = tupkey[0]
        index = tupkey[1]
        if oldkey in YN_list:
            newkey = tup[0]
            strYN = tup[1]
            if "TransactionInput.PricingInput.Pnr.Segments.IsForceStopover" == oldkey:
                useIsForceStopover.append(index)
            elif "TransactionInput.PricingInput.Pnr.Segments.IsForceConnection" == oldkey:
                if index in useIsForceStopover:
                    removeList.append(tup)
                    continue
                else:
                    newkey = "TransactionInput.PricingInput.Pnr.Segments[%d].IsForceStopover"%(index)
            modifyList.append((i,(newkey,ConverToTrueFalse(strYN))))

    #modify
    for i in range(0, len(modifyList)):
        oldJsonList[modifyList[i][0]] = modifyList[i][1]

    #remove
    for i in range(0, len(removeList)):
        oldJsonList.remove(removeList[i])

    return
#########################################################################

def processDiagnostic(oldJsonList, json_list_list):
    mlist = []
    listName = "TransactionInput.PricingInput.Options.Diagnostic"
    typeName = "TransactionInput.PricingInput.Options.Diagnostic.DiagnosticType"
    for i in range(0, len(oldJsonList)):
        tup = oldJsonList[i]
        oldkey = tup[0]
        if typeName in oldkey:
            if getDiagnosticContext(oldkey) != "":
                mlist.append(tup)

    for i in range(0, len(json_list_list)):
        if listName == json_list_list[i][0][0]:
            tup = (json_list_list[i][0],len(mlist))
            json_list_list.remove(json_list_list[i])
            json_list_list.append(tup)

    for i in range(0, len(mlist)):
        if i == 0:
            tup = ("TransactionInput.PricingInput.Options.Diagnostic.DiagnosticType", getDiagnosticContext(mlist[i][0]))
            oldJsonList.append(tup)
            tup = ("TransactionInput.PricingInput.Options.Diagnostic.DiagnosticInclude", ConverToTrueFalse(mlist[i][1]))
            oldJsonList.append(tup)
        else:
            tup = ("TransactionInput.PricingInput.Options.Diagnostic[%d].DiagnosticType"%(i+1), getDiagnosticContext(mlist[i][0]))
            oldJsonList.append(tup)
            tup = ("TransactionInput.PricingInput.Options.Diagnostic[%d].DiagnosticInclude"%(i+1), ConverToTrueFalse(mlist[i][1]))
            oldJsonList.append(tup)

        oldJsonList.remove(mlist[i])

    return
#########################################################################
def processSource(oldJsonList, json_list_list):
    agent_list = []
    mlist = []
    listName = "TransactionInput.PricingInput.Agent"
    for i in range(0, len(oldJsonList)):
        tup = oldJsonList[i]
        oldkey = tup[0]
        if listName in oldkey: 
            if getAgentContext(oldkey) != "":
                mlist.append(tup)
            else:
                agent_list.append(tup)

    for i in range(0, len(json_list_list)):
        if listName == json_list_list[i][0][0]:
            tup = (json_list_list[i][0],len(mlist))
            json_list_list.remove(json_list_list[i])
            json_list_list.append(tup)

    for i in range(0, len(mlist)):
        if i == 0:
            tup = ("TransactionInput.PricingInput.Agent.Request.ID", mlist[i][1])
            oldJsonList.append(tup)
            tup = ("TransactionInput.PricingInput.Agent.Request.ID_Context", getAgentContext(mlist[i][0]))
            oldJsonList.append(tup)
        else:
            tup = ("TransactionInput.PricingInput.Agent[%d].Request.ID"%(i+1), mlist[i][1])
            oldJsonList.append(tup)
            tup = ("TransactionInput.PricingInput.Agent[%d].Request.ID_Context"%(i+1), getAgentContext(mlist[i][0]))
            oldJsonList.append(tup)
            for j in range(0, len(agent_list)):
                
                keyold = agent_list[j][0]
                keyName = listName + "[%d]"%((i+1)) + keyold[len(listName):len(keyold)]
                tup = (keyName, agent_list[j][1])
                oldJsonList.append(tup)

        oldJsonList.remove(mlist[i])

    return
#########################################################################
def procDateTime(oldJsonList, key_old, tup, dicDateTime, index, prex, key1, key2):
    if key_old == prex + key1:
        dicDateTime.setdefault("Date",tup)
        if "Time" in list(dicDateTime.keys()) and len(dicDateTime["Time"]) != 0:
            if index > 1:
                key_old = prex +"[%d]"%(index) + key1 + "Time"
            else:
                key_old = prex + key1 + "Time"
            value = dicDateTime["Date"][1] + 'T' + dicDateTime["Time"][1]
            oldJsonList.append((key_old,value))
            dicDateTime.clear()
        else:
            return True
    elif key_old == prex + key2:
        dicDateTime.setdefault("Time",tup)
        if "Date" in list(dicDateTime.keys()) and len(dicDateTime["Date"])  != 0:
            if index > 1:
                key_old = prex +"[%d]"%(index) + key1 + "Time"
            else:
                key_old = prex + key1 + "Time"
            value = dicDateTime["Date"][1] + 'T' + dicDateTime["Time"][1]
            oldJsonList.append((key_old,value))
            dicDateTime.clear()
        else:
            return True
    return False
#########################################################################
def processDateTime(oldJsonList):
    dicDateTimeDep = {}
    DepDateTime = 0

    dicDateTimeArr = {}
    ArrDateTime = 0

    dicDateTimeRes = {}
    ResDateTime = 0

    prex = "TransactionInput.PricingInput.Pnr.Segments"
    depkey1 = ".DepDate"
    depkey2 = ".DepTime"
    arrkey1 = ".ArrDate"
    arrkey2 = ".ArrTime"
    reskey1 = ".ResDate"
    reskey2 = ".ResTime"

    for i in range(0, len(oldJsonList)):
        tup = oldJsonList[i]
        tupkey = getPurKey(tup[0])
        key_old = tupkey[0]
        index = tupkey[1]
        value = tup[1]

        if procDateTime(oldJsonList, key_old, tup, dicDateTimeDep, index, prex, depkey1, depkey2):
            continue
        if procDateTime(oldJsonList, key_old, tup, dicDateTimeArr, index, prex, arrkey1, arrkey2):
            continue
        if procDateTime(oldJsonList, key_old, tup, dicDateTimeRes, index, prex, reskey1, reskey2):
            continue
    return

#########################################################################
def readOldJson(fileName, list_list):
    fp = open(fileName, 'r')
    dict_json = json.loads(fp.read())
#    json.dump(dict_json, open('../Temp/rawjson.json', 'w'))

    key = ""
    oldJsonList = []
    json_list_list = []
    processDictionary(dict_json, key, list_list, oldJsonList, json_list_list, 0)

    processSource(oldJsonList, json_list_list)

    processDiagnostic(oldJsonList, json_list_list)

    processDateTime(oldJsonList)

    converYNToTrueFalse(oldJsonList)

    fp.close()
    return (oldJsonList,json_list_list)

#########################################################################

#########################################################################
def setValue(dict_json, keys, value, list_keys, list_index):
    idx = keys.find(".")
    key = ''
    subkeys = ''
    if idx == -1:
        key = keys
    else:
        key = keys[0:idx]
        subkeys = keys[idx+1:len(keys)]

    idx_list = list_keys.find(".")
    listkey = ''
    sublistkeys = ''
    if idx_list == -1:
        listkey = list_keys
    else:
        listkey = list_keys[0:idx_list]
        sublistkeys = list_keys[idx_list+1:len(list_keys)]

    index = 0
    if key == listkey:
        if sublistkeys == '':
            index = list_index
    else:
        sublistkeys = ''

    if key in list(dict_json.keys()):
        obj = dict_json[key]
        dic = OrderedDict()
        if str(type(obj)) == "<class 'collections.OrderedDict'>":
            dic = obj
        elif str(type(obj)) == "<class 'list'>":
            dic = obj[index]
        else:
            dict_json[key] = value
        
        if len(dic) != 0 and subkeys!= '':
            setValue(dic, subkeys, value, sublistkeys, list_index)

    return

#########################################################################

#########################################################################
def setList(dict_json, keys, count):
    index = keys.find(".")
    key = ''
    subkeys = ''
    if index != -1:
        key = keys[0:index]
        subkeys = keys[index+1:len(keys)]
    else:
        key = keys

    if key in list(dict_json.keys()):
        obj = dict_json[key]
        dic = OrderedDict()
        if str(type(obj)) == "<class 'collections.OrderedDict'>":
            dic = obj
        elif str(type(obj)) == "<class 'list'>":
            if subkeys == '':
                if len(obj) == count:
                    return
                else:
                    item = dict_json[key][0]
                    for i in range (1, count):
                        dict_json[key].append(copyItem(item))
            else:
                dic = obj[0]
        else:
            return
        
        if len(dic) != 0 and subkeys!= '':
            setList(dic, subkeys, count)

    return

#########################################################################

def processTemplate(new_Json_Template, json_list_list):
    new_Json = new_Json_Template
    for i in range(0,len(json_list_list)):
        count = json_list_list[i][1]
        if count > 1:
            keys = json_list_list[i][0][1]
            setList(new_Json, keys, count)
    return new_Json
    

#########################################################################
def convert(fileName , dict_element, list_list, new_Json_Template):
    ret = readOldJson(fileName, list_list)
    oldJsonList = ret[0]
    json_list_list = ret[1]
#    json.dump(oldJsonList, open('../Temp/oldjsonfile.json', 'w'))
#    json.dump(json_list_list, open('../Temp/oldjsonListfile.json', 'w'))

    new_Json = processTemplate(new_Json_Template, json_list_list)

#    json.dump(new_Json, open('../Temp/newjsontemplate.json', 'w'))

#    json.dump(list(dict_element.keys()), open('../Temp/dict_element_key_list.json', 'w'))

    list_element_keys = list(dict_element.keys())

    for i in range(0, len(oldJsonList)):
        tup = oldJsonList[i]
        key_old_raw = tup[0]
        value = tup[1]

        list_key = ""
        list_index = 0
        key_old = key_old_raw
        idx = key_old_raw.find("[")
        if idx != -1:
            idx2 = key_old_raw.find("]")
            list_index_str = key_old_raw[idx+1:idx2]
            if list_index_str.isnumeric():
                list_index = int(list_index_str)
            list_key = key_old_raw[0:idx]
            key_old = list_key + key_old_raw[idx2+1:len(key_old_raw)]

        new_list_key = ''
        if list_index > 1:
            for i in range(0, len(json_list_list)):
                if list_key == json_list_list[i][0][0]:
                    new_list_key = json_list_list[i][0][1]
                    break

        if key_old in list_element_keys:
            key_new = dict_element[key_old]
            setValue(new_Json, key_new, value, new_list_key, list_index-1)

    return new_Json

#########################################################################

def convertFiles(source, dest, new_Json_Template, dict_element, list_list):
    os.mkdir(dest)
    for root, dirs, files in os.walk( source ):
        for OneFileName in files :
            inputFileName = root + "/" + OneFileName
            outputFileName = dest + root[8:len(root)] + "/" + OneFileName
            new_Json_Temp = copyItem(new_Json_Template)
            new_Json = convert(inputFileName, dict_element, list_list, new_Json_Temp)

            json.dump(new_Json, open(outputFileName, 'w'))
            print ("%s is OK" %(OneFileName))
        if len(dirs) != 0:
            for dir in dirs:
                os.mkdir(dest + "/" + root[8:len(root)] + "/" + dir)
    return

#########################################################################
tempdir = "../Temp"
delete_file_folder(tempdir)
os.mkdir(tempdir)

fileName = "../data/dictionary.xls"
dict_element = OrderedDict()
list_list = []
getDictionary(fileName, dict_element, list_list)

#json.dump(list_list, open('../Temp/dictionarylistfile.json', 'w'))
#json.dump(dict_element, open('../Temp/dictionaryelementfile.json', 'w'))

new_Json_Template = readTemplate(fileName)
#json.dump(new_Json_Template, open('../Temp/jsontemplate.json', 'w'))

source = "../input"
dest = "../output"
delete_file_folder(dest)

convertFiles(source, dest, new_Json_Template, dict_element, list_list)

print ("All finished!")

