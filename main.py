from HYSYS_python_spreadsheets import Aspen_connection
import time
from ExcelManager import getDataFromExcel, saveParamsToExcel

# 1.0 Data of the Aspen HYSYS file
File         = 'Test_1.hsc'
Spreadsheets = ('SS_Flash', 'SS_turbine', 'SS_Distillation')
Units        = ('Cooler', 'Flash Drum', 'Heater', 'Valve', 'Reactor',
                'Distillation Column', 'Turbine', 'Pump')

# 2.0 Perform connection
workingFile      = Aspen_connection(File, Spreadsheets, Units)
# Turbine     = workingFile.SS['SS_turbine']
# Efficiency  = Turbine.Cell(1,0)            # .Cell(Column,Row) starting from 0
# Generation  = Turbine.Cell(1,1)
# ori_eff     = Efficiency.CellValue
solver      = workingFile.Solver


def measureEverethingWithExcel(excelName):
    allInputValues = getDataFromExcel(excelName)
    params = list(allInputValues.keys())
    allOutputValues = {} # ALL i.e. {"Температура":[2.0 9.0], "Давление":[2.0 9.0]}

    for i in range(len(allInputValues[params[0]])): # loop for all values in each param
        disposableInputDict = {} # {"T" : 0, "P": 1}
        for p in params:
            val = allInputValues[p][i]
            disposableInputDict[p] = val
        disposableMeasures = takeOneMeasureOfInputValues("P-Feed", disposableInputDict, "Feed")
        outputParams = list(disposableMeasures.keys())

        for p in outputParams:
            try:
                allOutputValues[p].append(disposableMeasures[p])
            except KeyError:
                allOutputValues[p] = [disposableMeasures[p]]

    saveParamsToExcel(excelName, allOutputValues)

def takeOneMeasureOfInputValues(inputStreamName, inputValues: dict, outputStreamName) -> dict: # inputValues = {"T": 9.0, "P": 9.0}
    solver.CanSolve = False  # Turn off the solving mode
    stream = workingFile.MaterialStreams[inputStreamName] # 'P-Feed'

    changeStream(stream, inputValues) # change our input values of stream
    solver.CanSolve = True  # Turn on the solving mode
    while solver.IsSolving == True: # waiting for Hysys to solve
        time.sleep(0.001)
    outputStream = workingFile.MaterialStreams[outputStreamName] # 'Feed'
    outputValues = readOutputStream(outputStream)
    return outputValues

def changeStream(stream, valuesDict: dict): # valuesDict = {"T" : 9.0, "P" : 6.7}
    for key in valuesDict:
        currentParamVal = valuesDict[key]
        if(key == "T"):
            stream.Temperature.value = currentParamVal
        elif(key == "P"):
            stream.Pressure.value = currentParamVal
        elif(key == "M"):
            stream.MolarFlow.value = currentParamVal

def readOutputStream(stream) -> dict: #{"температура": 76.0, "давление": 77.0"}
    result = {}
    result[stream.Temperature.Description] = stream.Temperature.Value
    result[stream.Pressure.Description] = stream.Pressure.Value
    result[stream.MolarFlow.Description] = stream.MolarFlow.Value
    return result


if __name__ == '__main__':
    measureEverethingWithExcel("test.xlsx")

