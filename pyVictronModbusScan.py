import os.path
import os
import time
import pandas
import requests
from pymodbus.constants import Endian
from pymodbus.client import ModbusTcpClient as ModbusClient
from pymodbus.payload import BinaryPayloadDecoder
import math
import numpy as np

colCharSize_index = 10
colCharSize_registerOverviewIndex = 10
colCharSize_registerDescription = 60
colCharSize_registerModbusAdr = 10
colCharSize_registerServiceName = 40
colCharSize_registerDbusObjPath = 60
colCharSize_value = 20
registerOverviewPagesEntryCount = 30
victronExcelFileHeaderRowIndexNumber = 1
baseStr = "modbus:\n  - name: victron\n    type: tcp\n    host: x.x.x.x\n    port: 502\n    sensors:"
fileURL = "https://raw.githubusercontent.com/victronenergy/dbus_modbustcp/master/CCGX-Modbus-TCP-register-list.xlsx"
fileTarget = "ModbusRegisterList.xlsx"
spaces = '      '

# Local network ip address of Cerbo GX. Default port 502
client = ModbusClient('XXX.XXX.XXX.XXX', port=502)

def cls(): os.system('cls' if os.name=='nt' else 'clear')

def getModbusExceptionStr(exceptionCode: int):
    if exceptionCode == 1:
        return "IllegalFunction"
    elif exceptionCode == 2:
        return "IllegalDataAddress"
    elif exceptionCode == 3:
        return "IllegalDataValue"
    elif exceptionCode == 4:
        return "GatewayPathUnavailable"
    elif exceptionCode == 5:
        return "GatewayTargetDeviceFailedToRespond"

def fillStringUpWithSpaces(inputStr: str, size: int) -> str:
    tmpStr = inputStr

    try:
        origSize = len(inputStr)
        fillUpAmount = size - origSize
    except:
        tmpStr = ""
        fillUpAmount = size   

    for x in range(fillUpAmount):
        tmpStr += " "

    return tmpStr

def parseExcelToDict(filepath: str, sheetName: str, headerRowNr: int):
    excel_data_df = pandas.read_excel(filepath, sheet_name = sheetName, header = headerRowNr)
    return excel_data_df.to_dict()

def getAllCellValuesFromColumn(xlsxDict: dict, colName: str):
    counter = 0
    tmpDict = {}
    for entryName in xlsxDict[colName].values():
        if entryName == "":
            tmpDict[counter] = "-"
        else:
            tmpDict[counter] = entryName
        counter += 1
    return tmpDict

def modbus_register_uint16(address, unit, factor: float):
    try:
        msg1     = client.read_holding_registers(address, slave=unit)
        decoder = BinaryPayloadDecoder.fromRegisters(msg1.registers, byteorder=Endian.BIG)
        msg2     = decoder.decode_16bit_uint() / factor
        return msg2
    except Exception as error:
        return "ERROR: " + str(getModbusExceptionStr(msg1.exception_code))

def modbus_register_uint32(address, unit, factor: float):
    try:
        msg1     = client.read_holding_registers(address, count=4, slave=unit)
        decoder = BinaryPayloadDecoder.fromRegisters(msg1.registers, byteorder=Endian.BIG)
        msg2     = decoder.decode_32bit_uint() / factor
        return msg2
    except Exception as error:
        return "ERROR: " + str(getModbusExceptionStr(msg1.exception_code))
        
def modbus_register_int16(address, unit, factor: float):
    try:
        msg1     = client.read_holding_registers(address, slave=unit)
        decoder = BinaryPayloadDecoder.fromRegisters(msg1.registers, byteorder=Endian.BIG)
        msg2     = decoder.decode_16bit_int() / factor
        return msg2
    except Exception as error:
        return "ERROR: " + str(getModbusExceptionStr(msg1.exception_code))

def modbus_register_int32(address, unit, factor: float):
    try:
        msg1     = client.read_holding_registers(address, count=4, slave=unit)
        decoder = BinaryPayloadDecoder.fromRegisters(msg1.registers, byteorder=Endian.BIG)
        msg2     = decoder.decode_32bit_int() / factor
        return msg2
    
    except Exception as error:
        return "ERROR: " + str(getModbusExceptionStr(msg1.exception_code))
        
def modbus_register_string(address, byteCount, unit):
    try:
        msg1     = client.read_holding_registers(address, byteCount, slave=unit)
        decoder = BinaryPayloadDecoder.fromRegisters(msg1.registers, byteorder=Endian.BIG)
        msg2     = decoder.decode_string(byteCount*2)
        return str(msg2).replace("\\x00", "")
    except Exception as error:
        return "ERROR" + str(error)

def strRemoveAllRepeatingWhitespaces(inputStr):
    return " ".join(inputStr.split())

def strGetSubstringBetween(buffer, strStart, strEnd):
    return buffer[buffer.find(strStart)+len(strStart):buffer.rfind(strEnd)]

def findDictKeyByValue(inutDict: dict, searchValue: str):
    for key, value in dict.items():
        if value == searchValue:
            return key
        
def createDictFromFile(filename: str):
    dictionary = {}
    with open(filename, 'r') as file:        
        for line in file:
            # Zeile in Schlüssel und Wert aufteilen
            key, value = line.strip().split(',')
            dictionary[key] = value
    return dictionary


def main():
    cls()
    print("----------------------------------------------------------------------------------------------------------------------------")
    print("Welcome to the Victron GX Modbus Register Converter for Homeassistant configuration.yaml")
    print("First we need the actual CCGX Modubus Register Definition file (xlsx file) to parse the available entries.")
    print("----------------------------------------------------------------------------------------------------------------------------")

    fileValidFlag = False
    inputStr = input("Do you want to load the latest CCGX Modbus register file from the Victron Github repository? [Y/n]: ")
    if inputStr.lower() == "y" or len(inputStr) == 0:
        try:
            r = requests.get(fileURL, allow_redirects=True)
            open(fileTarget, 'wb').write(r.content)
            filepath = "./" + fileTarget
            print("File download successful")
            fileValidFlag = True
        except:
            print("Something went wrong while loading the file, aborting the script.")
            exit(-2)
    else:
        print("Please get it and specify its location here:")        
        while fileValidFlag == False:
            filepath = input("File: ")
            if not os.path.isfile(filepath):
                print("Specified file does not exist, try again.")
            elif 'xlsx' not in filepath:
                print("Specified file is not a xlsx file, try again.")
            else:
                print("Ok, the specified file is valid")
                fileValidFlag = True

    dict_complete = parseExcelToDict(filepath, 'Field list', victronExcelFileHeaderRowIndexNumber)
    dict_names = getAllCellValuesFromColumn(dict_complete, 'description')
    dict_paths = getAllCellValuesFromColumn(dict_complete, 'dbus-obj-path')
    dict_modbusAdr = getAllCellValuesFromColumn(dict_complete, 'Address')
    dict_serviceName = getAllCellValuesFromColumn(dict_complete, 'dbus-service-name')
    dict_type = getAllCellValuesFromColumn(dict_complete, 'Type')
    dict_factor = getAllCellValuesFromColumn(dict_complete, 'Scalefactor')
    list_serviceNamesSelected = []
    countEntries = len(dict_names)


    print("I have found " + str(countEntries) + " registers in the list")

    serviceDict = createDictFromFile('myVictronModbusRegisters.txt')

    output = ""
    print(fillStringUpWithSpaces("Index:", colCharSize_index) +
          fillStringUpWithSpaces("Name:", colCharSize_registerServiceName) +
          fillStringUpWithSpaces("Description:", colCharSize_registerDescription) +
          fillStringUpWithSpaces("Adress:", colCharSize_registerModbusAdr) + 
          fillStringUpWithSpaces("ObjPath:", colCharSize_registerDbusObjPath) + 
          fillStringUpWithSpaces("Value:", colCharSize_value))
    print("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
    
    for register_index in range(0, countEntries):

        
        if(dict_serviceName[register_index] in serviceDict) and (isinstance(dict_names[register_index], str) and "RESERVED" not in dict_names[register_index]):    
            try:
                if strRemoveAllRepeatingWhitespaces(dict_type[register_index]) == "uint16":
                    output = str(modbus_register_uint16(int(dict_modbusAdr[register_index]), int(serviceDict[dict_serviceName[register_index]]), float(dict_factor[register_index])))
                elif strRemoveAllRepeatingWhitespaces(dict_type[register_index]) == "int16":
                    output = str(modbus_register_int16(int(dict_modbusAdr[register_index]), int(serviceDict[dict_serviceName[register_index]]), float(dict_factor[register_index])))
                elif strRemoveAllRepeatingWhitespaces(dict_type[register_index]) == "uint32":
                    output = str(modbus_register_uint32(int(dict_modbusAdr[register_index]), int(serviceDict[dict_serviceName[register_index]]), float(dict_factor[register_index])))
                elif strRemoveAllRepeatingWhitespaces(dict_type[register_index]) == "int32":
                    output = str(modbus_register_int32(int(dict_modbusAdr[register_index]), int(serviceDict[dict_serviceName[register_index]]), float(dict_factor[register_index])))
                elif "string" in strRemoveAllRepeatingWhitespaces(dict_type[register_index]):
                    strLen = strGetSubstringBetween(strRemoveAllRepeatingWhitespaces(dict_type[register_index]), "[", "]")
                    output = str(modbus_register_string(int(dict_modbusAdr[register_index]), int(strLen), int(serviceDict[dict_serviceName[register_index]])))

                print(fillStringUpWithSpaces("[" + str(register_index + 3) + "]", colCharSize_index) + 
                        fillStringUpWithSpaces(dict_serviceName[register_index], colCharSize_registerServiceName) + 
                        fillStringUpWithSpaces(dict_names[register_index], colCharSize_registerDescription) +
                        fillStringUpWithSpaces(str(dict_modbusAdr[register_index]), colCharSize_registerModbusAdr) + 
                        fillStringUpWithSpaces(dict_paths[register_index], colCharSize_registerDbusObjPath) + 
                        fillStringUpWithSpaces(output, colCharSize_value))
            
            except Exception as error:
                print("Error at Adr: " + str(dict_modbusAdr[register_index]) + " ", error)

            time.sleep(0.02)


if __name__ == "__main__":
    main()