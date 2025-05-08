#!/usr/bin/env python
import time
import pyautogui
import pygetwindow as gw
import mouse
#import autoMethods as am  
import datetime
#import sys 
#import os
#sys.path.append(os.path.abspath("C:/Laborsteuerung/MeasurementAutomation/"))
#from voltageControl_andSave_methods_h import *

#this program is providing the setup of automatic measurements in the lab of the High energy physics group of the chair of physik und ihre didaktik
#author is burkhard boehm

#definition of todays time
tday = datetime.date.today()
#Coordinates that are display and resolution depending
pauseTime = 0.6
selectDir_x = 328 
selectDir_y = 245
datLineAmptek_x=450
datLineAmptek_y=570
datLineExcel_x = 540
datLineExcel_y = 670
excelBrowse_x = 450
excelBrowse_y =  500

#################--------------------general methods----------------########################## 
#type in the name of the window you want to open ans state true if ypu want to maximize the window
def openAnyWindowAndMaximize(windowType, max=True):
    windowInsert = [window for window in gw.getAllTitles() if windowType in window]
    if windowInsert:
        # Bring the first Excel window to the foreground
        target_window_title = windowInsert[0]
        window = gw.getWindowsWithTitle(target_window_title)[0]
        if window:
            window.activate()
            #window.minimize()
            #window.maximize()minimizeActiveWindow
            if max == True:
                window.maximize()
            #window.activate()
        #print(f'Switched to the Window: {target_window_title}')
    else:
        print('No Window found with the name!')

def minimizeActiveWindow(windowType):
    windowInsert = [window for window in gw.getAllTitles() if windowType in window]
    if windowInsert:
        # Bring the first Excel window to the foreground
        target_window_title = windowInsert[0]
        window = gw.getWindowsWithTitle(target_window_title)[0]
        if window:
            #window.activate()
            window.minimize()
            #window.maximize()


#this method finds a button that looks like the picture and provides the coordinates
def findButtonCoordinates(pathToPic = 'OnScreenPicsLocal\stopLabview.png'):
    buttonCoord = pyautogui.locateCenterOnScreen(pathToPic, confidence=0.8)
    pyautogui.moveTo(buttonCoord[0], buttonCoord[1])
    buttonX = buttonCoord[0]
    buttonY = buttonCoord[1]
    return buttonX, buttonY

def createDir(measType, date, partsPerMillion, scanType, driftVolume ):
    #create a directory
    #time.sleep(pauseTime*2)
    dir = measType + str(date) + "_" + partsPerMillion + scanType + driftVolume
    print(dir)
    time.sleep(pauseTime)
    openAnyWindowAndMaximize("Amptek DppMCA", True)
    time.sleep(pauseTime*4)
    #dirButton = findButtonCoordinates('OnScreenPicsLocal\saveDir.png')
    #pyautogui.moveTo(dirButton[0], dirButton[1])
    #pyautogui.click()
    pyautogui.hotkey('ctrl','o')
    time.sleep(pauseTime*4)
    pyautogui.press('F4')
    time.sleep(pauseTime)
    pyautogui.hotkey('ctrl','a')
    #fw = pyautogui.getActiveWindow()
    #fw.topleft = (200, 200)
    #time.sleep(pauseTime)
    #pyautogui.moveTo(selectDir_x, selectDir_y)
    #pyautogui.click()
    time.sleep(pauseTime)
    #create new directory for the current measurement
    pyautogui.press('backspace')
    time.sleep(pauseTime)
    pyautogui.write('C:/data')
    time.sleep(pauseTime)
    pyautogui.press('enter')
    time.sleep(pauseTime)
    pyautogui.hotkey('ctrl','shift','n')
    time.sleep(pauseTime*2)
    pyautogui.write(dir)
    time.sleep(pauseTime*2)
    pyautogui.press('enter')
    time.sleep(pauseTime)
    pyautogui.hotkey('Alt', 'F4')
    time.sleep(pauseTime)
    return dir 

################----------------methods for the controle of CAENHV Labview--------------##################################################
#check if it is already running 
def isRunning():
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    if pyautogui.locateOnScreen('OnScreenPicsLocal\greenLight.png', confidence=0.98):
         print("The Programm to run the CAEN Power Supply is running correctly.")
         return True
    else:
        print("Code is not running correctly!!!!")
        return False 

#stopLabview
def stopLabview():
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    if(isRunning()):
        pyautogui.hotkey('ctrl','.')

#start_Labview()
def startLabview():
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    if(isRunning()):
        #do nothing
        print("Programm is stopped, so code not running is normal.")
        time.sleep(pauseTime)
    else:
        pyautogui.hotkey('ctrl','r')
        time.sleep(pauseTime)
        buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\startCAEN.png', confidence=0.73)
        #ipLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\ipAddress.png', confidence=0.73)
        #pyautogui.moveTo(ipLocation[0], ipLocation[1]+25)
        #pyautogui.click()
        #pyautogui.click()
        #pyautogui.press('backspace')
        #time.sleep(pauseTime)
        #insert ip-address of the power supply
        #pyautogui.write("132.187.196.33")
        pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
        pyautogui.click()
        

def executeCommand(command):
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\execute.png', confidence=0.75)        
    pyautogui.moveTo(buttonLocation[0]-96, buttonLocation[1]-2)
    pyautogui.click()
    time.sleep(pauseTime)
    if command == "Kill" or command == "kill":
        pyautogui.press("down")
        pyautogui.press("enter")
    elif command == "ClearAlarm" or command == "clearalarm" or command == "Clearalarm":     
        pyautogui.press("down")
        pyautogui.press("down")
        pyautogui.press("enter")
    pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
    pyautogui.click()

def selectChannelWindow():
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\saveChannel.png', confidence=0.75)
    pyautogui.moveTo(buttonLocation[0]-187, buttonLocation[1]-3)
    pyautogui.click()

def setChannelValue(chVal):
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\saveChannel.png', confidence=0.75)
    pyautogui.moveTo(buttonLocation[0]-140, buttonLocation[1]+2)
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(pauseTime)
    pyautogui.write(chVal)
    pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
    pyautogui.click()

############################## switch On MM channels ##############################################
def switchOnMMChannels(meshVolt = "460", driftVolt = "702"):
    #prepare labview for bugfree change
    stopLabview(), startLabview()
    #start channel 10 Mesh
    loadParameterList("5","10")
    #set Voltage for the mesh/channel 10
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(1):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue(meshVolt)
    #move to I_0 and put values at 2.0 nA for ramping up
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(1):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("2,00")
    #explicitly start the channel
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("1")
    #refreshON()
    
    #start channel 9 anode
    loadParameterList("5","9")
    #set Voltage for the mesh/channel 9 
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(11):
        pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("0,00")
    #move to I_0 and put values at 2.0 nA for ramping up
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(1):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("2,00")
    #explicitly start the channel
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("1")
    #refreshON()

    #start channel 8 drift
    loadParameterList("5","8")
    #set Voltage for the mesh/channel 8 
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(11):
        pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue(driftVolt)
    #move to I_0 and put values at 2.0 nA for ramping up
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(1):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("2,00")
    #explicitly start the channel
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("1")
    #refreshON()
    #wait and set I_0_max to 0.2 nA for each channel
    time.sleep(60)#usually 250nA for testign 50nA
    
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("0,2")
    loadParameterList("5","9")
    time.sleep(pauseTime)
    setChannelValue("0,2")
    loadParameterList("5","10")
    time.sleep(pauseTime)
    setChannelValue("0,2")
    #refreshON()

############################## switch Off MM channels ##############################################
def switchOffMMChannels():
    pauseTime = 0.5
    stopLabview(), startLabview()
    #switch off channel 10 Mesh
    loadParameterList("5","10")
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(2):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("2,0")
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("0")
    #refrshOn()
    #witch off channel 9 anode
    loadParameterList("5","9")
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("2,0")
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("0")
    #refreshON()
    #witch off channel 8 drift
    loadParameterList("5","8")
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("2,0")
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(10):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("0")
    #refreshON()

#to set a parameter one wants to change to a certain value for a channel providing the channel number    
def setParameterOfChannel(parameterToChange, parameterValue, channelNumber):
    #prepare labview for bugfree change
    #stopLabview(), startLabview()
    openAnyWindowAndMaximize("CAEN", True)
    time.sleep(1)
    loadParameterList("5",channelNumber)
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(30):
            pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    if parameterToChange == "V_0" or parameterToChange == "V0" or parameterToChange == "Voltage" or parameterToChange == "voltage":
        for i in range(1):
            pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(pauseTime)
        setChannelValue(parameterValue)
    elif parameterToChange == "I_0" or parameterToChange == "I0" or parameterToChange == "Current" or parameterToChange == "current":
        for i in range(2):
            pyautogui.press('down')
        pyautogui.press('enter')
        time.sleep(pauseTime)
        setChannelValue(parameterValue)
    channelNumber = "default"
    channel = "default"
    #refreshON()

def switchOnSingleChannel(channelNumber):
    #prepare labview for bugfree change
    #stopLabview(), startLabview()
    loadParameterList("5",channelNumber)
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(30):
        pyautogui.press('up')
    time.sleep(pauseTime)
    
    for i in range(12):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("1")
    #refreshON()

def switchOFFSingleChannel(channelNumber):
    #prepare labview for bugfree change
    #stopLabview(), startLabview()
    loadParameterList("5",channelNumber)
    selectChannelWindow()
    time.sleep(pauseTime)
    for i in range(30):
        pyautogui.press('up')
    time.sleep(pauseTime)
    for i in range(12):
        pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(pauseTime)
    setChannelValue("0")
    #refreshON()

def loadParameterList(slot, channel):
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\loadParameterList.png', confidence=0.8)
    pyautogui.moveTo(buttonLocation[0]-230, buttonLocation[1])
    time.sleep(pauseTime)
    #print(buttonLocation)
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(pauseTime)
    pyautogui.write(slot)
    pyautogui.moveTo(buttonLocation[0]-150, buttonLocation[1])
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(pauseTime)
    pyautogui.write(channel)
    pyautogui.moveTo(buttonLocation[0]-10, buttonLocation[1])
    pyautogui.click()

def refreshON():
    if pyautogui.locateOnScreen('OnScreenPicsLocal\OFFrefresh.png', confidence=0.8):
        buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\OFFrefresh.png', confidence=0.8)
        #print(buttonLocation)
        pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
        pyautogui.click()
        pyautogui.move(60, 0)

def refreshOFF():
    if pyautogui.locateOnScreen('OnScreenPicsLocal\ONrefresh.png', confidence=0.8):
        buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\ONrefresh.png', confidence=0.8)
        #print(buttonLocation)
        pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
        pyautogui.click()
        pyautogui.move(60, 0)
    

def testButtons():
    if pyautogui.locateOnScreen('OnScreenPics\OFFrefresh.png', confidence=0.8):
        pyautogui.moveTo('OnScreenPics\OFFrefresh.png')
        buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPics\OFFrefresh.png', confidence=0.8)
        print(buttonLocation)

def setMeshVoltage(vmm = "460"):
    timeNowMesh = datetime.datetime.now()
    openAnyWindowAndMaximize("CAENHVWrapper",True)
    #set Mesh voltage 
    setParameterOfChannel("V_0", vmm, "10")
    print(timeNowMesh.strftime("%d-%m-%Y %Xh") + ": Mesh Voltage was set to " + vmm)

def setDriftVoltage(vd = "702"):
    timeNowDrift = datetime.datetime.now()
    openAnyWindowAndMaximize("CAENHVWrapper",True)
    #set Mesh voltage 
    setParameterOfChannel("V_0", vd, "8")
    print(timeNowDrift.strftime("%d-%m-%Y %Xh") + ": Drift Voltage was set to " + vd)

def setCurrent(curThresh = "0,2"):
    #caen_window = [window for window in gw.getAllTitles() if "CAEN" in window]
    openAnyWindowAndMaximize("CAENHVWrapper",True)
    #set Mesh voltage 
    setParameterOfChannel("I_0", curThresh, "10")
    #set Drift voltage 
    setParameterOfChannel("I_0", curThresh, "8")
    setParameterOfChannel("I_0", curThresh, "9")

def doAGainScan(timeForMeas = 180, driftVolume = "5mm", measurementType = "H2O", ppm = "0ppm", todaysDate = "2001-12-20", save = True, maxVolt = 500):
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan was started")
    #create a directory
    scanT = "_ampScan"
    #time.sleep(pauseTime)
    if(save):
        dir = createDir(measurementType, todaysDate, ppm, scanT, driftVolume)
   
    if(driftVolume == "5mm"):
        if(maxVolt == 500):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 500VA was started")
            vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan for 5mm drift volume
            vec_driftVolt = ["672","687","702","717","733","748","763"]
            print("500 is the maximum mesh Voltage")
        elif(maxVolt == 480):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 480VA was started")
            vec_meshVolt = ["440","450","460","470","480"] #define the mesh voltage for a gain scan for 5mm drift volume
            vec_driftVolt = ["672","687","702","717","733"]    
            print("480 is the maximum mesh Voltage")
        elif(maxVolt == 520):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 520VA was started")
            vec_meshVolt = ["460","470","480","490","500","510","520"] #define the mesh voltage for a gain scan for 5mm drift volume
            vec_driftVolt = ["702","717","733","748","763","778","793"]
            print("500 is the maximum mesh Voltage")
    elif(driftVolume == "4mm"):
        if(maxVolt == 500):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 500VA was started")
            vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan for 4mm drift volume
            vec_driftVolt = ["626","640","654","668","682","696","710"]
            print("500 is the maximum mesh Voltage")
        elif(maxVolt == 480):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 480VA was started")
            vec_meshVolt = ["440","450","460","470","480"] #define the mesh voltage for a gain scan for 4mm drift volume
            vec_driftVolt = ["626","640","654","668","682"]
            print("480 is the maximum mesh Voltage")
        elif(maxVolt == 520):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 520VA was started")
            vec_meshVolt = ["460","470","480","490","500","510","520"] #define the mesh voltage for a gain scan for 4mm drift volume
            vec_driftVolt = ["654","668","682","696","710","724","738"]
            print("500 is the maximum mesh Voltage")
    elif(driftVolume == "3mm"):
        if(maxVolt == 500):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 500VA was started")
            vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan for 3mm drift volume
            vec_driftVolt = ["579","592","605","618","632","645","658"]
            print("500 is the maximum mesh Voltage")
        elif(maxVolt == 480):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 480VA was started")
            vec_meshVolt = ["440","450","460","470","480"] #define the mesh voltage for a gain scan for 3mm drift volume
            vec_driftVolt = ["579","592","605","618","632"]
            print("480 is the maximum mesh Voltage")
        elif(maxVolt == 520):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 520VA was started")
            vec_meshVolt = ["460","470","480","490","500","510","520"] #define the mesh voltage for a gain scan for 3mm drift volume
            vec_driftVolt = ["605","618","632","645","658","671","684"]
            print("500 is the maximum mesh Voltage")
    elif(driftVolume == "2.5mm"):
        if(maxVolt == 500):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 500VA was started")
            vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan for 2.5mm drift volume
            vec_driftVolt = ["556","569","581","594","607","619","632"]
            print("500 is the maximum mesh Voltage")
        elif(maxVolt == 480):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 480VA was started")
            vec_meshVolt = ["440","450","460","470","480"] #define the mesh voltage for a gain scan for 2.5mm drift volume
            vec_driftVolt = ["556","569","581","594","607"]
            print("480 is the maximum mesh Voltage")
        elif(maxVolt == 520):
            print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 520VA was started")
            vec_meshVolt = ["460","470","480","490","500","510","520"] #define the mesh voltage for a gain scan for 2.5mm drift volume
            vec_driftVolt = ["581","594","607","619","632","644","657"]
            print("500 is the maximum mesh Voltage")
    #vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan
    #vec_driftVolt = ["672","687","702","717","733","748","763"]
    #vec_meshVolt = ["440","450"] #define the mesh voltage for a gain scan
    #vec_driftVolt = ["672","687"]
    #openAnyWindowAndMaximize("CAENHVWrapper",True)
    index = len(vec_meshVolt)
    #measurement upwards
    time.sleep(pauseTime)
    for i in range(index):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setMeshVoltage(vec_meshVolt[i])
        time.sleep(pauseTime)
        setDriftVoltage(vec_driftVolt[i])
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime*2)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = vec_meshVolt[i]+"VA"+vec_driftVolt[i]+"VD"  
        if(save):
            saveSpectrum(dir, file) 
        #time.sleep(10)

    #measurement downwards, doing a reverse for-loop
    time.sleep(pauseTime)
    for i in range(index-1,-1,-1):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setMeshVoltage(vec_meshVolt[i])
        time.sleep(pauseTime)
        setDriftVoltage(vec_driftVolt[i])
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime*2)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = vec_meshVolt[i]+"VA"+vec_driftVolt[i]+"VD"+"_2"  
        if(save):
            saveSpectrum(dir, file)     
        #time.sleep(10)

def doAGainScanFromCertainValue(timeForMeas = 180, driftVolume = "5mm", measurementType = "H2O", ppm = "0ppm", todaysDate = "2001-12-20"):
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan was started")
    #create a directory
    scanT = "_ampScan"
    dir = measurementType + "-" + str(todaysDate) + "_" + ppm + scanT
    vec_meshVolt = ["440","450","460","470"] #define the mesh voltage for a gain scan
    vec_driftVolt = ["672","687","702","717"]
    #vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan
    #vec_driftVolt = ["672","687","702","717","733","748","763"]
    #openAnyWindowAndMaximize("CAENHVWrapper",True)
    index = len(vec_meshVolt)
    #measurement upwards
  
    
    #measurement downwards, doing a reverse for-loop
    time.sleep(pauseTime)
    for i in range(index-1,-1,-1):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setMeshVoltage(vec_meshVolt[i])
        time.sleep(pauseTime)
        setDriftVoltage(vec_driftVolt[i])
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = vec_meshVolt[i]+"VA"+vec_driftVolt[i]+"VD"+"_2"  
        saveSpectrum(dir, file)     
        #time.sleep(10)

#now the pressure and teh electric field are changing
def doAGainScanWithDependingElectricField(timeForMeas = 180, driftVolume = "5mm", measurementType = "E-Field_H2O", ppm = "0ppm", todaysDate = "2001-12-20", save = True, p_value = 1010):
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": E-Field depending Gain Scan was started")
    #create a directory
    scanT = "_ampScan"
    #time.sleep(pauseTime)
    if(save):
        dir = createDir(measurementType, todaysDate, ppm, scanT, driftVolume)
   
    p_std = 1010 
  
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gain Scan 500VA was started")
    vec_meshVolt = [440,450,460,470,480,490,500] #define the mesh voltage for a gain scan
    vec_driftVolt = [672,687,702,717,733,748,763]
    index = len(vec_meshVolt)
    
    p_factor = p_value/p_std
    vec_meshVolt_facAdded = [round(p_factor * element) for element in vec_meshVolt]
    vec_driftVolt_facAdded = [round(p_factor * element) for element in vec_driftVolt]
    for i in range(index):
       print("Meas. all voltages: " , vec_meshVolt_facAdded[i], vec_driftVolt_facAdded[i])
   
    #vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan
    #vec_driftVolt = ["672","687","702","717","733","748","763"]
    #vec_meshVolt = ["440","450"] #define the mesh voltage for a gain scan
    #vec_driftVolt = ["672","687"]
    #openAnyWindowAndMaximize("CAENHVWrapper",True)
    #measurement upwards
    time.sleep(pauseTime)
    for i in range(index):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setMeshVoltage(str(vec_meshVolt_facAdded[i]))
        time.sleep(pauseTime)
        setDriftVoltage(str(vec_driftVolt_facAdded[i]))
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)
        
        #save the MCA file as .mcs and .csv file
        file = str(vec_meshVolt_facAdded[i])+"VA"+str(vec_driftVolt_facAdded[i])+"VD"  
        if(save):
            saveSpectrum(dir, file)
        #time.sleep(10)

    #measurement downwards, doing a reverse for-loop
    time.sleep(pauseTime)
    for i in range(index-1,-1,-1):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setMeshVoltage(str(vec_meshVolt_facAdded[i]))
        time.sleep(pauseTime)
        setDriftVoltage(str(vec_driftVolt_facAdded[i]))
        refreshON()
        time.sleep(pauseTime)
        #makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = str(vec_meshVolt_facAdded[i])+"VA"+str(vec_driftVolt_facAdded[i])+"VD"+"_2"  
        if(save):
            saveSpectrum(dir, file)  
        #time.sleep(10)
    
def doNoFullGainScanWithDependingElectricField(timeForMeas = 180, driftVolume = "5mm", measurementType = "E-Field_H2O", ppm = "0ppm", todaysDate = "2001-12-20", save = True, p_value = 1010):
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": E-Field depending no full Gain Scan was started")
    #create a directory
    scanT = "_ampScan"
    #time.sleep(pauseTime)
    if(save):
        dir = createDir(measurementType, todaysDate, ppm, scanT, driftVolume)
   
    p_std = 1010 
  
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Short Gain Scan 450VA/490VA was started")
    vec_meshVolt = [450,490] #define the mesh voltage for a gain scan
    vec_driftVolt = [687,748]
    index = len(vec_meshVolt)
    
    p_factor = p_value/p_std
    vec_meshVolt_facAdded = [round(p_factor * element) for element in vec_meshVolt]
    vec_driftVolt_facAdded = [round(p_factor * element) for element in vec_driftVolt]
    for i in range(index):
        print("Meas short voltages: " , vec_meshVolt_facAdded[i], vec_driftVolt_facAdded[i])
   
    #vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan
    #vec_driftVolt = ["672","687","702","717","733","748","763"]
    #vec_meshVolt = ["440","450"] #define the mesh voltage for a gain scan
    #vec_driftVolt = ["672","687"]
    #openAnyWindowAndMaximize("CAENHVWrapper",True)
    #measurement upwards
    time.sleep(pauseTime)
    for i in range(index):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setMeshVoltage(str(vec_meshVolt_facAdded[i]))
        time.sleep(pauseTime)
        setDriftVoltage(str(vec_driftVolt_facAdded[i]))
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = str(vec_meshVolt_facAdded[i])+"VA"+str(vec_driftVolt_facAdded[i])+"VD"  
        if(save):
            saveSpectrum(dir, file) 
        #time.sleep(10)

    #measurement downwards, doing a reverse for-loop
    time.sleep(pauseTime)
    for i in range(index-1,-1,-1):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setMeshVoltage(str(vec_meshVolt_facAdded[i]))
        time.sleep(pauseTime)
        setDriftVoltage(str(vec_driftVolt_facAdded[i]))
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = str(vec_meshVolt_facAdded[i])+"VA"+str(vec_driftVolt_facAdded[i])+"VD"+"_2"  
        if(save):
            saveSpectrum(dir, file)     
        #time.sleep(10)

        
def doADriftScan(meshVoltage = "500",driftVolume = "5mm", timeForMeas = 180, measurementType = "H2O", ppm = "0ppm", todaysDate = "2001-12-20", save = True):
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Drift Scan was started")
    #create a directory
    scanT = "_driftScan"
    time.sleep(pauseTime)
    if(save):
        dir = createDir(measurementType, todaysDate, ppm, scanT, driftVolume)
    openAnyWindowAndMaximize("CAENHVWrapper", True)
 

    #set mesh voltage for the drift scan
    setMeshVoltage(meshVoltage)
    vec_driftVolt = []
    if (driftVolume == "5mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["550","600","650","700","750","800","850","900","950","1000","1050","1100"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["530","580","630","680","730","780","830","880","930","980","1030","1080"]
    elif (driftVolume == "4mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["540","580","620","660","700","740","780","820","860","900","940","980"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["520","560","600","640","680","720","760","800","840","880","920","960"]
    elif (driftVolume == "3mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["530","560","590","620","650","680","710","740","770","800","830","880"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["510","540","570","600","630","660","690","720","750","780","810","840"]
    elif (driftVolume == "2.5mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["525","550","575","600","625","650","675","700","725","750","775","800"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["505","530","555","580","605","630","655","680","705","730","755","780"]
    #vec_driftVolt = ["650","700","750","800","850","900","950","1000","1050","1100"]
    #vec_driftVolt = ["650","700"]
    index = len(vec_driftVolt)

    #measurement upwards
    time.sleep(pauseTime)
    for i in range(index):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setDriftVoltage(vec_driftVolt[i])
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = meshVoltage + "VA" + vec_driftVolt[i] + "VD"  
        if(save):
            saveSpectrum(dir, file) 
        #time.sleep(10)

    time.sleep(pauseTime)
   
    #measure the zero of the gain measurement again
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    setMeshVoltage("440")
    time.sleep(pauseTime)
    setDriftVoltage("672")
    refreshON()
    time.sleep(pauseTime)
    makeSpectrum(timeForMeas)
    time.sleep(pauseTime)
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    refreshOFF()
    time.sleep(pauseTime)

    #save the MCA file as .mcs and .csv file
    file = "440VA672VD"  
    if(save):
        saveSpectrum(dir, file) 
    #time.sleep(10)

    #measurement downwards, doing a reverse for-loop
    time.sleep(pauseTime)
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    setMeshVoltage(meshVoltage)
    time.sleep(10)

    for i in range(index-1,-1,-1):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setDriftVoltage(vec_driftVolt[i])
        if(vec_driftVolt[i]=="1080"):
            time.sleep(120)
            #print("Test")
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = meshVoltage + "VA" + vec_driftVolt[i] + "VD"+"_2"  
        if(save):
            saveSpectrum(dir, file)     
        #time.sleep(10)

    #measure the zero of the gain measurement again
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    setMeshVoltage("440")
    time.sleep(pauseTime)
    setDriftVoltage("672")
    refreshON()
    time.sleep(pauseTime)
    makeSpectrum(timeForMeas)
    time.sleep(pauseTime)
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    refreshOFF()
    time.sleep(pauseTime)

    #save the MCA file as .mcs and .csv file
    file = "440VA672VD_2"  
    if(save):
        saveSpectrum(dir, file) 
    #time.sleep(10)

def doADriftScanFromCertainValue(meshVoltage = "500",driftVolume = "5mm", timeForMeas = 180, dir_input = "Comb_0_2_2024-04-02_O2_4000ppm_H2O_0ppm_driftScan", todaysDate = "2001-12-20", save = True):
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Drift Scan was started")
    #create a directory
    scanT = "_driftScan"
    time.sleep(pauseTime)
    dir = dir_input
    openAnyWindowAndMaximize("CAENHVWrapper", True)
 

    #set mesh voltage for the drift scan
    setMeshVoltage(meshVoltage)
    vec_driftVolt = []
    if (driftVolume == "5mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["550","600","650","700","750","800","850","900","950","1000","1050","1100"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["530","580","630","680","730","780","830","880","930","980","1030","1080"]
    elif (driftVolume == "4mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["540","580","620","660","700","740","780","820","860","900","940","980"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["520","560","600","640","680","720","760","800","840","880","920","960"]
    elif (driftVolume == "3mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["530","560","590","620","650","680","710","740","770","800","830","880"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["510","540","570","600","630","660","690","720","750","780","810","840"]
    elif (driftVolume == "2.5mm"):
        if(meshVoltage == "500"):
            vec_driftVolt = ["525","550","575","600","625","650","675","700","725","750","775","800"]
        elif(meshVoltage == "480"):
            vec_driftVolt = ["505","530","555","580","605","630","655","680","705","730","755","780"]
    #vec_driftVolt = ["650","700","750","800","850","900","950","1000","1050","1100"]
    #vec_driftVolt = ["650","700"]
    index = len(vec_driftVolt)
    '''
    #measurement upwards
    time.sleep(pauseTime)
    for i in range(index):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setDriftVoltage(vec_driftVolt[i])
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = meshVoltage + "VA" + vec_driftVolt[i] + "VD"  
        if(save):
            saveSpectrum(dir, file) 
        #time.sleep(10)

    time.sleep(pauseTime)
'''   
    #measure the zero of the gain measurement again
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    setMeshVoltage("440")
    time.sleep(pauseTime)
    setDriftVoltage("672")
    refreshON()
    time.sleep(pauseTime)
    makeSpectrum(timeForMeas)
    time.sleep(pauseTime)
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    refreshOFF()
    time.sleep(pauseTime)

    #save the MCA file as .mcs and .csv file
    file = "440VA672VD"  
    if(save):
        saveSpectrum(dir, file) 
    #time.sleep(10)

    #measurement downwards, doing a reverse for-loop
    time.sleep(pauseTime)
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    setMeshVoltage(meshVoltage)
    time.sleep(10)
    
    for i in range(index-1,-1,-1):
        #change voltages for mesh and drift channels
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        setDriftVoltage(vec_driftVolt[i])
        if(vec_driftVolt[i]=="1080"):
            time.sleep(120)
            #print("Test")
        time.sleep(pauseTime)
        refreshON()
        time.sleep(pauseTime)
        makeSpectrum(timeForMeas)
        time.sleep(pauseTime)
        openAnyWindowAndMaximize("CAENHVWrapper", True)
        time.sleep(pauseTime)
        refreshOFF()
        time.sleep(pauseTime)

        #save the MCA file as .mcs and .csv file
        file = meshVoltage + "VA" + vec_driftVolt[i] + "VD"+"_2"  
        if(save):
            saveSpectrum(dir, file)     
        #time.sleep(10)

    #measure the zero of the gain measurement again
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    setMeshVoltage("440")
    time.sleep(pauseTime)
    setDriftVoltage("672")
    refreshON()
    time.sleep(pauseTime)
    makeSpectrum(timeForMeas)
    time.sleep(pauseTime)
    openAnyWindowAndMaximize("CAENHVWrapper", True)
    time.sleep(pauseTime)
    refreshOFF()
    time.sleep(pauseTime)

    #save the MCA file as .mcs and .csv file
    file = "440VA672VD_2"  
    if(save):
        saveSpectrum(dir, file) 
    #time.sleep(10)
        
# this method is for doing measurements to check how accurate a measurement represents several other measurments taht are done later        
def checkAccuracyOfMeas(timeForMeas = 300, measurementType = "Acc", ppm = "0ppm", todaysDate = "2001-12-20"):
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Accuracy measurement was started")
    #create a directory
    dir = createDir(measurementType, todaysDate, ppm, "_errorScan", "5mm")
    
    
    vec_meshVolt = ["440","450","460","470","480","490","500"] #define the mesh voltage for a gain scan
    vec_driftVolt = ["672","687","702","717","733","748","763"]
    #openAnyWindowAndMaximize("CAENHVWrapper",True)
    index = len(vec_meshVolt)

    time.sleep(pauseTime)
    for j in range(0,20):
    #for j in range(0,1):
        for i in range(index):
            #change voltages for mesh and drift channels
            openAnyWindowAndMaximize("CAENHVWrapper", True)
            time.sleep(pauseTime)
            setMeshVoltage(vec_meshVolt[i])
            time.sleep(pauseTime)
            setDriftVoltage(vec_driftVolt[i])
            refreshON()
            time.sleep(10)
            makeSpectrum(timeForMeas)
            time.sleep(pauseTime)
            openAnyWindowAndMaximize("CAENHVWrapper", True)
            time.sleep(pauseTime)
            refreshOFF()
            time.sleep(pauseTime*2)

            #save the MCA file as .mcs and .csv file
            file = vec_meshVolt[i]+"VA"+vec_driftVolt[i]+"VD"+"_"+ str(j)
            saveSpectrum(dir, file) 
            time.sleep(3)
        time.sleep(300)
    #time.sleep(6)
    
   

###########----------------------------------- Amptek DppMCA and excel methods -------------#######################
#measure a spectrum with the AMptek DppMCA program
def makeSpectrum(specTime):
    openAnyWindowAndMaximize("Amptek DppMCA", True)
    time.sleep(pauseTime*3)
    pyautogui.press('a')
    #buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\delData.png', confidence=0.7)
    #pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
    #pyautogui.click()
    time.sleep(pauseTime*3)
    
    pyautogui.press('F3')
    '''
    #start_measurement()
    if pyautogui.locateOnScreen('OnScreenPicsLocal\start.png', confidence = 0.8):
        #buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\start.png', confidence=0.7)
        #pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
        #pyautogui.click()
        time.sleep(pauseTime)
        pyautogui.press('F3')
    else: 
        timeNow2 = datetime.datetime.now()
        print(timeNow2.strftime("%d-%m-%Y %Xh") + ": Measurement is still running! No start needed.")
        ()
    '''
    time.sleep(specTime)
    #pyautogui.move(0, 40)
    pyautogui.press('F3')
    #buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\stop.png', confidence=0.8)
    #pyautogui.moveTo(buttonLocation[0], buttonLocation[1])
    #pyautogui.click()
    time.sleep(pauseTime*3)
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Spectrum measurement was done sucessfully.")

#change the gas with LabView
def changeGasContent(ppmO2, ppmH2O):
    time.sleep(2)
    openAnyWindowAndMaximize("SC_V4",False)
    time.sleep(pauseTime*2)
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\setO2.png', confidence=0.7)
    pyautogui.moveTo(buttonLocation[0], buttonLocation[1]+30)
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(0.2)
    pyautogui.write(ppmO2)
    pyautogui.press('enter')
    

    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\setH2O.png', confidence=0.7)
    pyautogui.moveTo(buttonLocation[0], buttonLocation[1]+30)
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(0.2)
    pyautogui.write(ppmH2O)
    pyautogui.press('enter')
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gas was changed to " + ppmO2 + "ppm O2 and " + ppmH2O + "ppm H2O." )

#change the gas with LabView
def changeGasFlow(flowAr):
    time.sleep(2)
    openAnyWindowAndMaximize("SC_V4",False)
    time.sleep(pauseTime*2)
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\setFlowAr.png', confidence=0.7)
    pyautogui.moveTo(buttonLocation[0], buttonLocation[1]+38)
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(0.2)
    pyautogui.write(flowAr)
    pyautogui.press('enter')
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gas flow was changed to " + flowAr + "ml/min." )

#change the pressure with LabView
def changeOverPressure(pressureChange):
    time.sleep(2)
    openAnyWindowAndMaximize("SC_V4",False)
    time.sleep(pauseTime*2)
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\setFlowAr.png', confidence=0.7)
    pyautogui.moveTo(buttonLocation[0]+845, buttonLocation[1]-250)
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    #time.sleep(0.2)
    pyautogui.write(str(pressureChange))
    pyautogui.press('enter')
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Ideal Pressure was changed to: " + str(pressureChange))

'''
def changeGasFlow(flowAr):
    time.sleep(2)
    openAnyWindowAndMaximize("SC_V4",False)
    time.sleep(pauseTime*2)
    buttonLocation = pyautogui.locateCenterOnScreen('OnScreenPicsLocal\setFlowAr.png', confidence=0.7)
    pyautogui.moveTo(buttonLocation[0], buttonLocation[1]+38)
    pyautogui.click()
    pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(0.2)
    pyautogui.write(flowAr)
    pyautogui.press('enter')
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Gas flow was changed to " + flowAr + "ml/min." )
'''

#save the spectrum measured by the DppMCA program
def saveSpectrum(dirName, fileName):
    #savin as .mcs
    openAnyWindowAndMaximize("Amptek DppMCA", True)
    time.sleep(pauseTime)
    pyautogui.hotkey('ctrl','s') # or use pyautogui.click('OnScreenPics\save.png')
    time.sleep(pauseTime)
    pyautogui.press('F4')
    time.sleep(pauseTime)
    pyautogui.hotkey('ctrl','a')
    time.sleep(pauseTime)
    #fw = pyautogui.getActiveWindow()
    #fw.topleft = (200, 200)
    #pyautogui.moveTo(selectDir_x, selectDir_y)
    #pyautogui.click()
    pyautogui.press('backspace')
    time.sleep(pauseTime)
    pyautogui.write('C:\data\\'+dirName)
    pyautogui.press('enter')
    #pyautogui.moveTo(datLineAmptek_x, datLineAmptek_y)
    #pyautogui.click()
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(pauseTime)
    pyautogui.press('backspace')
    pyautogui.write(fileName+'.mcs')
    pyautogui.press('enter')
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Spectrum saved successfully as .mcs")
    time.sleep(pauseTime*4)
    
    #the save as csv part is not used anymore because there were to many bugs causing the program to fail and measurements had to be restarted too often
    ''' 
    #now save as .csv    
    openAnyWindowAndMaximize("Amptek DppMCA", True)
    pyautogui.hotkey('ctrl','r')
    time.sleep(pauseTime*4)
    
    #openAnyWindowAndMaximize("Excel", False)
    windowInsert = [window for window in gw.getAllTitles() if "Excel" in window]
    if windowInsert:
        # Bring the first Excel window to the foreground
        target_window_title = windowInsert[0]

        window = gw.getWindowsWithTitle(target_window_title)[0]
        if window:
            window.activate()
    else: #alternative if now excel window is open, open a new one
        print('No Window found with the name!')
        pyautogui.hotkey('ctrl','esc')
        time.sleep(0.5)
        pyautogui.write("Excel")
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(5)
        pyautogui.hotkey('Alt','F4')
        time.sleep(0.2)
        pyautogui.press('enter')

    time.sleep(pauseTime)
    #get to excel window and insert the data and save the file
    pyautogui.hotkey('ctrl','v')
    time.sleep(pauseTime)
    pyautogui.hotkey('ctrl','s')
    time.sleep(pauseTime)
    fw = pyautogui.getActiveWindow()
    fw.topleft = (200, 200)
    pyautogui.moveTo(excelBrowse_x, excelBrowse_y)
    pyautogui.click()
  
    time.sleep(pauseTime)
    pyautogui.press('F4')
    time.sleep(pauseTime)
    pyautogui.hotkey('ctrl','a')
    time.sleep(pauseTime)
    pyautogui.press('backspace')
    time.sleep(pauseTime)
    pyautogui.write('C:\data\\'+dirName)
    time.sleep(pauseTime)
    pyautogui.press('enter')
    time.sleep(pauseTime)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(pauseTime)
    pyautogui.press('backspace')
    time.sleep(pauseTime)
   
    pyautogui.write(fileName)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.hotkey('ctrl','n')

    time.sleep(3)
    
    openAnyWindowAndMaximize(fileName+".csv", False)
    time.sleep(pauseTime)
    #fw.topleft = (200, 200)
    pyautogui.hotkey('Alt','F4')
    time.sleep(pauseTime)
    timeNow = datetime.datetime.now()
    print(timeNow.strftime("%d-%m-%Y %Xh") + ": Spectrum saved successfully as .csv")
'''

#main to execute all the methods shown above
def main():
    #measurement e-field and pressure change at different water concentrations
    vec_p_ideal = [985,990,995,1000,1005,1010,1015,1020,1025,1030,1035,1040,1045,1050,1055,1060] 
    vec_H2O_pre = [0, 16000, 3000, 0]
    vec_H2O =     [0, 12000, 6000, 1]
    
    counting_p = range(len(vec_p_ideal))
    counting_h2o = range(len(vec_H2O))
    
    doAGainScanWithDependingElectricField(240,"5mm","Test", "test", tday, True, 1010)

    changeGasFlow("100")
    changeGasContent("0","0")
    for j in counting_h2o:
        time.sleep(pauseTime)   
         
        print("Pre-water value change loop is running.")
        changeGasContent("0",str(vec_H2O_pre[j]))
        time.sleep(3600*1.5)
        print("Normal measurement loop is running.")
        changeGasContent("0",str(vec_H2O[j]))
        time.sleep(3600*3)
        time.sleep(pauseTime)
        
        for i in counting_p:
            changeOverPressure(vec_p_ideal[i])
            time.sleep(3600*0.1)
            dirName = "PressTest_" + str(i) + "_" + str(vec_p_ideal[i]) + "mbar_"
            gasVar = "H2O_" + str(vec_H2O[j]) + "ppm"
            time.sleep(pauseTime*2)
            print(gasVar)
            time.sleep(pauseTime*2)
            driftVol = "5mm"
                
            if(vec_p_ideal[i] == 985 or vec_p_ideal[i]== 1010 or vec_p_ideal[i] == 1035 or vec_p_ideal[i]== 1060 ):
                #print("fulfilled", i)
                doAGainScanWithDependingElectricField(240,driftVol,dirName, gasVar, tday, True, vec_p_ideal[i])
            else:
                #print("if condition not fulfilled")
                doNoFullGainScanWithDependingElectricField(240,driftVol, dirName, gasVar, tday, True, vec_p_ideal[i])

    changeGasFlow("40")




    #measurements for water and oxygen measurement
    vec_meshVolt = ["440","450","460","470","480","490","500","490","480","470","460","450","440"] #define the mesh voltage for a gain scan
    vec_driftVolt = ["672","687","702","717","733","748","763","748","733","717","702","687","672"]
    #vec_H2O_pre = [0,18000,2000,24000,2000,24000,4000,0]
    #vec_H2O =     [0,9000,6000,12000,6000,12000,9000,0]
    #vec_H2O_pre = [0,3000,6000,8000,10000,12000,0]
    #vec_H2O =     [0,2000,3000,4000,5000,6000,0]
    #vec_H2O_pre = [0,15000,1000,22000,6000,24000]
    #vec_H2O =     [0,8000,4000,16000,12000,18000]
    #stats 30.9 3mm measure
    #standard H2O for measurements
    #vec_H2O_pre = [0, 10000, 4000, 16000]
    #vec_H2O =     [0, 8000, 6000, 12000]    
    #stats 30.9 3mm measure
    #standard O2 for measurements
    vec_O2_pre = [0,4000,100,6000,8000,400,8000,0]
    vec_O2 = [0,2000,500,4000,8000,1000,6000,0]
    #vec_O2_pre = [0,3000,6000,0]
    #vec_O2 = [0,2000,4000,0]
    #vec_O2_pre = [0]
    #vec_O2 = [0]
    vec = [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]
    
   
    '''
    for i in vec:
        time.sleep(pauseTime)
        dirName = "Acc"+ str(i) + "_"
        time.sleep(pauseTime)
        doAGainScan(240,dirName, "O2_0ppm_H2O_0ppm", tday, True, 500)  #duration ~1h15min
        time.sleep(2*3600)                                           #duration 30min total = 1h45min
    #start of 1st water Oxygen measurement block H2O = 1500
    '''
    counting = range(len(vec_O2))
    counting_2 = range(len(vec_H2O_pre))
    #dirName = "InAndOutTest_2_"
    #gasVar = "O2_0ppm_H2O_0ppm"
    #driftVol = "5mm"
            
    '''
    for i in counting:
        print("Pre-Oxygen value change loop is running.")
        changeGasContent(str(vec_O2_pre[i]),"0")
        time.sleep(3600*0.5)
        print("Normal measurement loop is running.")
        time.sleep(pauseTime)
        changeGasContent(str(vec_O2[i]),"0")
        time.sleep(3600*1)   
    '''
    '''time.sleep(3600*2)
    #changeGasFlow("45")
    #doAGainScan(120,driftVol,dirName, gasVar, tday, True,500)  #duration ~52min
    changeGasFlow("45")
    for j in counting_2:
        #time.sleep(pauseTime)
        print("Pre-water value change loop is running.")
        #if(j>0):
        changeGasContent("0",str(vec_H2O_pre[j]))
        time.sleep(3600*1)
        #time.sleep(10)
        for i in counting:  
            print("Pre-Oxygen value change loop is running.")
            changeGasContent(str(vec_O2_pre[i]),str(vec_H2O[j]))
            time.sleep(3600*0.5)
            if(vec_O2_pre[i]==8000):
                   time.sleep(3600*0.5) 
            print("Normal measurement loop is running.")
            time.sleep(pauseTime)
            changeGasContent(str(vec_O2[i]),str(vec_H2O[j]))
            time.sleep(3600*1.5)   
            #time.sleep(10)   
            #print("Set O2: " + str(vec_O2[i]) ,"Set H2O: " + str(vec_H2O[j]))
            time.sleep(pauseTime)
            dirName = "Comb_" + str(j) + "_" + str(i) + "_"
            gasVar = "O2_" + str(vec_O2[i]) + "ppm_H2O_" + str(vec_H2O[j]) + "ppm"
            time.sleep(pauseTime*2)
            print(gasVar)
            time.sleep(pauseTime*2)
            #time.sleep(3600*1.0)
            driftVol = "2.5mm"
            doAGainScan(120,driftVol,dirName, gasVar, tday, True,500)  #duration ~52min
            doADriftScan("480",driftVol,120,dirName, gasVar, tday, True)#duration ~1h15min
    changeGasFlow("0")
    '''


    '''
            if(j<1 and i<4):
                print("Gas was not changed and measurement was skipped cause already done before.")
            elif(j==0 and i==4):
                time.sleep(pauseTime)
                dirName = "Comb_" + str(j) + "_" + str(i) + "_"
                gasVar = "O2_" + str(vec_O2[i]) + "ppm_H2O_" + str(vec_H2O[j]) + "ppm"
                time.sleep(pauseTime*2)
                print(gasVar)
                time.sleep(pauseTime*2)
                doAGainScan(120,dirName, gasVar, tday, True,500)  #duration ~52min
                #doADriftScan("480",60,dirName, gasVar, tday, True)#duration ~1h15min
                doADriftScan("480",120,dirName, gasVar, tday, True)#duration ~1h15min
            elif(j==0 and i>4):
                print("Pre-Oxygen value change loop is running.")
                changeGasContent(str(vec_O2_pre[i]),str(vec_H2O[j]))
                time.sleep(3600*0.5)
                print("Normal measurement loop is running.")
                time.sleep(pauseTime)
                changeGasContent(str(vec_O2[i]),str(vec_H2O[j]))
                time.sleep(3600*1)   
                #time.sleep(10)   
                #print("Set O2: " + str(vec_O2[i]) ,"Set H2O: " + str(vec_H2O[j]))
                time.sleep(pauseTime)
                dirName = "Comb_" + str(j) + "_" + str(i) + "_"
                gasVar = "O2_" + str(vec_O2[i]) + "ppm_H2O_" + str(vec_H2O[j]) + "ppm"
                time.sleep(pauseTime*2)
                print(gasVar)
                time.sleep(pauseTime*2)
                doAGainScan(120,dirName, gasVar, tday, True,500)  #duration ~52min
                #doADriftScan("480",60,dirName, gasVar, tday, True)#duration ~1h15min
                doADriftScan("480",120,dirName, gasVar, tday, True)#duration ~1h15min
            else:
            '''
    '''
    for j in counting_2:
        #time.sleep(pauseTime)
        print("Water value change loop is running.")
        changeGasContent("0",str(vec_H2O[j]))
        time.sleep(3600*1)
        #time.sleep(10)
        for i in counting:
            print("Normal measurement loop is running.")
            time.sleep(pauseTime)
            changeGasContent(str(vec_O2[i]),str(vec_H2O[j]))
            time.sleep(3600*1.5)   
            time.sleep(pauseTime)
            dirName = "Comb_" + str(j) + "_" + str(i) + "_"
            gasVar = "O2_" + str(vec_O2[i]) + "ppm_H2O_" + str(vec_H2O[j]) + "ppm"
            time.sleep(pauseTime*2)
            print(gasVar)
            time.sleep(3600*1)
    '''        
              
    #doing test runs taken a fast test run
    #time.sleep(3600*4)
    #doAGainScan(120,"GainTest_500_", "O2_17ppm_H2O_280ppm", tday, True, 500)  #duration ~52min
    #doAGainScan(120,"GainTest_520_", "O2_17ppm_H2O_280ppm", tday, True, 520)  #duration ~52min
    #doADriftScan("500",120,"SourcePos_1_Test_16mm_480_", "O2_0ppm_H2O_0ppm", tday, True)#duration ~1h15min       
    #doAGainScan(120,"SourcePosTest_16mm_480_", "O2_0ppm_H2O_0ppm", tday, True, 500)  #duration ~52min
    #doADriftScan("480",120,"SourcePosTest_16mm_480_", "O2_0ppm_H2O_0ppm", tday, True)#duration ~1h15min       
    
    


if __name__ == "__main__":
    main()

