from tkinter import ttk
from tkinter import *
import xlwings as xw
import numpy as np
import SerialDataGateway
import json
import re

class excelBridge:
    def __init__(self,wsheet):
        self._currentWorksheet = wsheet
        self._originRange = None
        self._variableNamesdictionary = None
        self._plotRowDict = None
        self._SelectedCell = {'row':-1,'col':-1}
        self._selectedRange = None
        self.currentPlot = None
        self.currentVariable = None
    def setupColumns(self):
        self._originRange = self._currentWorksheet.book.selection
        if self._originRange.count == 1:
            if len(re.findall('plot$',self._originRange.value,re.IGNORECASE)) != 0:
                lastUsedColumn = self._currentWorksheet.range(self._originRange.row,1000).end('left')
                variablenamesRange = self._currentWorksheet.range(self._originRange,lastUsedColumn)
                variableColumns = np.array(range(0,variablenamesRange.count)) + variablenamesRange.column
                self._variableNamesdictionary = {k:v for k,v in zip(variablenamesRange.value,variableColumns) if k is not None}
                self._SelectedCell['row'] = variablenamesRange.row
                return(list(self._variableNamesdictionary.keys()))
        else:
            return("bad")
    def setFocus(self):
        if self.currentPlot != "":
            try:
                plotRows = np.array( range(0,self._originRange.expand('down').count) ) + self._originRange.row
                self._plotRowDict = {str(k):v for k,v in zip(self._originRange.expand('down').options(numbers = int).value,plotRows)}
                self._SelectedCell['row'] = self._plotRowDict[self.currentPlot]
            except:
                print("Plot no encontrado")
                self._selectedRange = None
                self._originRange.select()
        if self.currentVariable != "":
            self._SelectedCell["col"] = self._variableNamesdictionary[self.currentVariable]
            self._selectedRange = self._currentWorksheet.range((self._SelectedCell['row'],self._SelectedCell['col']))
            self._selectedRange.select()

    def writeValue(self,value):
        if self._selectedRange is not None and self.currentPlot != "":
            self._selectedRange.value = value


class mainApp:
    def __init__(self,window):
        self.wind = window
        self.wind.title('Scale data collector')
        self.wind.resizable(0,0)

        self._devicesFile = "devices.json"
        self._devicesDict = None
        self._Status = False #Connection status
        self._CurrentPlotnum = None
        self._CurrentScaleValue = None
        self._CurrentWB = None
        # objects for devices
        self._ScaleGateway = None
        self._ScannerGateway = None

        #Excel and COM ports finder conectionFrame container
        conectionFrame = LabelFrame(self.wind, text = "Taget file and devices")
        conectionFrame.grid(row = 0, column = 0, pady = 5, padx = 10) 
        statusFrame  = LabelFrame(self.wind, text = "Data status")
        statusFrame.grid(row = 1, column = 0, pady = 5, padx = 10) 

        #Excel files opened combobox
        self._cOpenedFiles = StringVar()
        self._cVariableList = StringVar()
        self._entryPlot = StringVar()
        self._entryScale = StringVar()
        self._entryScanner = StringVar()
        self._entryValue = StringVar()

        #self._cOpenedFiles.trace('w',self.ConnectToWorksheet)
        self.combo = ttk.Combobox(conectionFrame,textvariable = self._cOpenedFiles,width = 45)
        self.combo.grid(row = 0, column = 0,sticky = W+E)
        self._ScaleTxEntry = ttk.Entry(conectionFrame,textvariable = self._entryScale)
        self._ScaleTxEntry.grid(row = 1, column = 0,sticky = W+E)
        self._ScannerTxEntry = ttk.Entry(conectionFrame,textvariable = self._entryScanner)
        self._ScannerTxEntry.grid(row = 2, column = 0,sticky = W+E)
        #Button to refresh the files list
        
        #ttk.Button(conectionFrame, text = 'Connect Scale').grid(row = 1,column = 1,sticky = W+E,columnspan = 2)
        #ttk.Button(conectionFrame, text = 'Connect Scanner').grid(row = 2,column = 2,sticky = W+E)

        Label(statusFrame,text = 'Variable').grid(row=0,column =0)
        self._VariableCombo = ttk.Combobox(statusFrame,textvariable = self._cVariableList,width = 35)
        self._VariableCombo.grid(row =0,column = 1, sticky = W+E)
        self._VariableCombo.bind("<<ComboboxSelected>>",self._SelectCell)

        Label(statusFrame,text = 'Plot').grid(row=1,column =0)
        self._PlotEntry = ttk.Entry(statusFrame,textvariable = self._entryPlot)
        self._PlotEntry.grid(row =1,column = 1,sticky = W)
        self._PlotEntry.bind("<Return>",self._SelectCell)
        self._entryPlot.trace('w',self._SelectCell)

        Label(statusFrame,text = 'Value').grid(row=2,column =0)
        self._ValueEntry = ttk.Entry(statusFrame, textvariable = self._entryValue)
        self._ValueEntry.grid(row =2,column = 1,sticky = W)
        self._entryValue.trace('w',self._WriteToWs)

        #ttk.Button(statusFrame,text = 'SendValue').grid(row = 3, column = 1)
        self._connectBttn = ttk.Button(statusFrame,text = 'Connect', command = self.ConnectToWorksheet)
        self._connectBttn.grid(row = 3, column = 0)
        ttk.Button(statusFrame, text = 'Refresh List', command = self.RefreshFilesList).grid(row = 3,column = 1 ,sticky = E)

    def RefreshFilesList(self):
        """
        Show the openned excel files and look for scale or barcode scanner devices 
        if not, do nothing    
        """
        # Refresh the list of excel opened files
        openedFiles = []
        for a in xw.apps:
            for w in a.books:
                openedFiles.append(w.name)
        self.combo['values'] = openedFiles
        self._cOpenedFiles.set("")
        if len(openedFiles) >= 1:
            self._cOpenedFiles.set(openedFiles[0])
        #Refresh the list of devices if nothing is connected
        if not self._Status:
            with open(self._devicesFile) as devf:
                self._devicesDict = json.load(devf)
                entryScale,entryScanner = SerialDataGateway.LookForDevices(self._devicesDict)
            self._entryScale.set(entryScale)
            self._entryScanner.set(entryScanner)

    def scaleLineHandler(self,line):
        self._CurrentScaleValue = "".join(re.findall('\d*\.*\d*',line))
        self._entryValue.set(self._CurrentScaleValue)
        #self._ValueEntry.delete(0,END)
        #self._ValueEntry.insert(0,self._CurrentScaleValue)
        print(line)
        
    def scannerLineHandler(self,line):
        self._CurrentPlotnum = line.split("_")[-2]
        self._entryPlot.set(self._CurrentPlotnum)
        #self._PlotEntry.delete(0,END)
        #self._PlotEntry.insert(0,self._CurrentPlotnum)
        print(line)

    def connectToDevices(self):
        #Check valid dict key
        if len(self._entryScale.get().split(',')) == 3:
            if self._entryScale.get().split(',')[-1] in self._devicesDict["Scale"].keys():
                scaleDevice = self._devicesDict["Scale"][self._entryScale.get().split(',')[-1]]
                self._ScaleGateway = SerialDataGateway.SerialDataGateway(port = self._entryScale.get().split(',')[1], baudrate=scaleDevice["baud"],bytesSize=scaleDevice["bits"],stopBits=scaleDevice["stopBits"],lineHandler=self.scaleLineHandler)
                self._ScaleGateway.Start()
        else:
            pass
        if len(self._entryScanner.get().split(',')) == 3:
            if self._entryScanner.get().split(',')[-1] in self._devicesDict["Scanner"].keys():
                scannerDevice = self._devicesDict["Scanner"][self._entryScanner.get().split(',')[-1]]
                self._ScannerGateway = SerialDataGateway.SerialDataGateway(port = self._entryScanner.get().split(',')[1], baudrate=scannerDevice["baud"],bytesSize=scannerDevice["bits"],stopBits=scannerDevice["stopBits"],lineHandler=self.scannerLineHandler)
                self._ScannerGateway.Start()
        else:
            pass

    def ConnectToWorksheet(self,*args):
        #Disable Entry devices
        if not self._Status: # if not connected
            self._ScaleTxEntry.config(state = DISABLED)
            self._ScannerTxEntry.config(state = DISABLED)
            self.combo.config(state = DISABLED)
            self.connectToDevices()
            try:
                self._CurrentWB = excelBridge(xw.books[self._cOpenedFiles.get()].sheets.active)
                self._VariableCombo['values'] = self._CurrentWB.setupColumns()
            except:
                pass
            self._connectBttn.config(text = 'Disconnect')
        else:
            self._ScaleTxEntry.config(state = NORMAL)
            self._ScannerTxEntry.config(state = NORMAL)
            self.combo.config(state = NORMAL)
            if self._ScaleGateway is not None:
                self._ScaleGateway.Stop()
            if self._ScannerGateway is not None:
                self._ScannerGateway.Stop()
            self._connectBttn.config(text = 'Connect')
            self._entryPlot.set("")
            self._entryValue.set("")
            self._VariableCombo['values'] = []
            self._cVariableList.set("")
            #delete worksheet bridge
        self._Status = not self._Status
    
    def _SelectCell(self,*args):
        if self._CurrentWB is not None:
            self._CurrentWB.currentVariable = self._cVariableList.get()
            self._CurrentWB.currentPlot = self._entryPlot.get()
            self._CurrentWB.setFocus()
        
    def _WriteToWs(self,*args):
        if self._CurrentWB is not None:
            self._CurrentWB.writeValue(self._entryValue.get())

if __name__ == '__main__':
    window = Tk()
    application = mainApp(window)
    window.mainloop()

