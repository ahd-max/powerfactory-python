import os
import sys
sys.path.append(r'D:\PowerFactory2017\Python\3.6')
import powerfactory as pf
app=pf.GetApplication()
import random
import numpy as np
import pandas as pd
from pandas import DataFrame
import math
app.ClearOutputWindow()
user=app.GetCurrentUser()

def setupSimultation(comInc, comSim, start_time, stop_time, period):
  comInc.iopt_sim = "rms"
  comInc.iopt_show = 0
  comInc.iopt_adapt = 0
  comInc.start = start_time
  comSim.tstop = stop_time
  comInc.dtgrd = period
##start_time=-0.1, stop_time=5/10, period=0.01

##clear last fault_event
def clearSimEvents():
  faultFolder = app.GetFromStudyCase("Simulation Events/Fault.IntEvt")
  cont = faultFolder.GetContents()
  for obj in cont:
    obj.Delete()

###add the fault event/clear the event
def addShcEvent(obj, sec, faultType):
  
  faultFolder = app.GetFromStudyCase("Simulation Events/Fault.IntEvt")
  event = faultFolder.CreateObject("EvtShc", obj.loc_name)
  event.p_target = obj
  event.time = sec
  event.i_shc = faultType

####load scalling
def loadscalling(loadscalling):
  load_list = app.GetCalcRelevantObjects('*.ElmLod')
  scale = loadscalling
  for oLoad in load_list :
    oLoad.scale0 = scale
    
def runSimulation(comInc, comSim):
  app.EchoOff()
  comInc.Execute()
  app.EchoOn()
  comSim.Execute()

def getResults(loc_fl,sec,loca):    
    #get result file
    elmRes = app.GetFromStudyCase('*.ElmRes')    
    app.ResLoadData(elmRes)
    
    #Get number of rows and columns
    NrRow = app.ResGetValueCount(elmRes,0)
    app.PrintPlain(NrRow)
     
    #get objects of interest
    buses = app.GetCalcRelevantObjects('*.ElmTerm')
    lines = app.GetCalcRelevantObjects('*.ElmLne')
    SG = app.GetCalcRelevantObjects('*.ElmSym')
    load_list = app.GetCalcRelevantObjects('*.ElmLod')

    c= {}
    for oLoad in load_list :
        c=oLoad.scale0
    
    LSG={}       
    for i in range(len(SG)):
        LSG[i]=SG[i].loc_name
    LSG=list(LSG.values())

    LLB={}       
    for i in range(len(buses)):
        LLB[i]=buses[i].loc_name
    #app.PrintPlain(LSG)
    LLBU=list(LLB.values())
    
    LLB=[str(i)+'_V' for i in LLBU]
    LLA=[str(i)+'_A' for i in LLBU]
    
    

    ColIndex_time = app.ResGetIndex(elmRes,elmRes,'b:tnow')
    ColIndex_angle1=app.ResGetIndex(elmRes,SG[0],'c:firel')
    ColIndex_angle2=app.ResGetIndex(elmRes,SG[1],'c:firel')
    ColIndex_angle3=app.ResGetIndex(elmRes,SG[2],'c:firel')
    ColIndex_angle4=app.ResGetIndex(elmRes,SG[3],'c:firel')
    ColIndex_angle5=app.ResGetIndex(elmRes,SG[4],'c:firel')
    ColIndex_angle6=app.ResGetIndex(elmRes,SG[5],'c:firel')
    ColIndex_angle7=app.ResGetIndex(elmRes,SG[6],'c:firel')
    ColIndex_angle8=app.ResGetIndex(elmRes,SG[7],'c:firel')
    ColIndex_angle9=app.ResGetIndex(elmRes,SG[8],'c:firel')
    ColIndex_angle10=app.ResGetIndex(elmRes,SG[9],'c:firel')

    ColIndex_speed1=app.ResGetIndex(elmRes,SG[0],'s:xspeed')
    ColIndex_speed2=app.ResGetIndex(elmRes,SG[1],'s:xspeed')
    ColIndex_speed3=app.ResGetIndex(elmRes,SG[2],'s:xspeed')
    ColIndex_speed4=app.ResGetIndex(elmRes,SG[3],'s:xspeed')
    ColIndex_speed5=app.ResGetIndex(elmRes,SG[4],'s:xspeed')
    ColIndex_speed6=app.ResGetIndex(elmRes,SG[5],'s:xspeed')
    ColIndex_speed7=app.ResGetIndex(elmRes,SG[6],'s:xspeed')
    ColIndex_speed8=app.ResGetIndex(elmRes,SG[7],'s:xspeed')
    ColIndex_speed9=app.ResGetIndex(elmRes,SG[8],'s:xspeed')
    ColIndex_speed10=app.ResGetIndex(elmRes,SG[9],'s:xspeed')

    ColIndex_u_bus1 = app.ResGetIndex(elmRes,buses[0],'m:u1')
    ColIndex_u_bus2 = app.ResGetIndex(elmRes,buses[1],'m:u1')
    ColIndex_u_bus3 = app.ResGetIndex(elmRes,buses[2],'m:u1')
    ColIndex_u_bus4 = app.ResGetIndex(elmRes,buses[3],'m:u1')
    ColIndex_u_bus5 = app.ResGetIndex(elmRes,buses[4],'m:u1')
    ColIndex_u_bus6 = app.ResGetIndex(elmRes,buses[5],'m:u1')
    ColIndex_u_bus7 = app.ResGetIndex(elmRes,buses[6],'m:u1')
    ColIndex_u_bus8 = app.ResGetIndex(elmRes,buses[7],'m:u1')
    ColIndex_u_bus9 = app.ResGetIndex(elmRes,buses[8],'m:u1')
    ColIndex_u_bus10 = app.ResGetIndex(elmRes,buses[9],'m:u1')
    ColIndex_u_bus11 = app.ResGetIndex(elmRes,buses[10],'m:u1')
    ColIndex_u_bus12 = app.ResGetIndex(elmRes,buses[11],'m:u1')
    ColIndex_u_bus13 = app.ResGetIndex(elmRes,buses[12],'m:u1')
    ColIndex_u_bus14 = app.ResGetIndex(elmRes,buses[13],'m:u1')
    ColIndex_u_bus15 = app.ResGetIndex(elmRes,buses[14],'m:u1')
    ColIndex_u_bus16 = app.ResGetIndex(elmRes,buses[15],'m:u1')
    ColIndex_u_bus17 = app.ResGetIndex(elmRes,buses[16],'m:u1')
    ColIndex_u_bus18 = app.ResGetIndex(elmRes,buses[17],'m:u1')
    ColIndex_u_bus19 = app.ResGetIndex(elmRes,buses[18],'m:u1')
    ColIndex_u_bus20 = app.ResGetIndex(elmRes,buses[19],'m:u1')
    ColIndex_u_bus21 = app.ResGetIndex(elmRes,buses[20],'m:u1')
    ColIndex_u_bus22= app.ResGetIndex(elmRes,buses[21],'m:u1')
    ColIndex_u_bus23 = app.ResGetIndex(elmRes,buses[22],'m:u1')
    ColIndex_u_bus24 = app.ResGetIndex(elmRes,buses[23],'m:u1')
    ColIndex_u_bus25 = app.ResGetIndex(elmRes,buses[24],'m:u1')
    ColIndex_u_bus26 = app.ResGetIndex(elmRes,buses[25],'m:u1')
    ColIndex_u_bus27 = app.ResGetIndex(elmRes,buses[26],'m:u1')
    ColIndex_u_bus28 = app.ResGetIndex(elmRes,buses[27],'m:u1')
    ColIndex_u_bus29 = app.ResGetIndex(elmRes,buses[28],'m:u1')
    ColIndex_u_bus30 = app.ResGetIndex(elmRes,buses[29],'m:u1')
    ColIndex_u_bus31 = app.ResGetIndex(elmRes,buses[30],'m:u1')
    ColIndex_u_bus32 = app.ResGetIndex(elmRes,buses[31],'m:u1')
    ColIndex_u_bus33 = app.ResGetIndex(elmRes,buses[32],'m:u1')
    ColIndex_u_bus34 = app.ResGetIndex(elmRes,buses[33],'m:u1')
    ColIndex_u_bus35 = app.ResGetIndex(elmRes,buses[34],'m:u1')
    ColIndex_u_bus36 = app.ResGetIndex(elmRes,buses[35],'m:u1')
    ColIndex_u_bus37 = app.ResGetIndex(elmRes,buses[36],'m:u1')
    ColIndex_u_bus38 = app.ResGetIndex(elmRes,buses[37],'m:u1')
    ColIndex_u_bus39 = app.ResGetIndex(elmRes,buses[38],'m:u1')

    ColIndex_u_busa1 = app.ResGetIndex(elmRes,buses[0],'m:phiurel')
    ColIndex_u_busa2 = app.ResGetIndex(elmRes,buses[1],'m:phiurel')
    ColIndex_u_busa3 = app.ResGetIndex(elmRes,buses[2],'m:phiurel')
    ColIndex_u_busa4 = app.ResGetIndex(elmRes,buses[3],'m:phiurel')
    ColIndex_u_busa5 = app.ResGetIndex(elmRes,buses[4],'m:phiurel')
    

    #pre-allocate result variables
    result_time = np.zeros((NrRow))
    result_name = np.zeros((NrRow),dtype=object)
    result_scalle = np.zeros((NrRow))
    result_ct = np.zeros((NrRow))
    result_loca = np.zeros((NrRow))
    
    ##rotor angle
    result_angle1 = np.zeros((NrRow))
    result_angle2 = np.zeros((NrRow))
    result_angle3 = np.zeros((NrRow))
    result_angle4 = np.zeros((NrRow))
    result_angle5 = np.zeros((NrRow))
    result_angle6 = np.zeros((NrRow))
    result_angle7 = np.zeros((NrRow))
    result_angle8 = np.zeros((NrRow))
    result_angle9 = np.zeros((NrRow))
    result_angle10 = np.zeros((NrRow))

    ##G speed
    result_speed1 = np.zeros((NrRow))
    result_speed2 = np.zeros((NrRow))
    result_speed3 = np.zeros((NrRow))
    result_speed4 = np.zeros((NrRow))
    result_speed5 = np.zeros((NrRow))
    result_speed6 = np.zeros((NrRow))
    result_speed7 = np.zeros((NrRow))
    result_speed8 = np.zeros((NrRow))
    result_speed9 = np.zeros((NrRow))
    result_speed10 = np.zeros((NrRow))

    #####bus v
    result_busv1 = np.zeros((NrRow))
    result_busv2 = np.zeros((NrRow))
    result_busv3 = np.zeros((NrRow))
    result_busv4 = np.zeros((NrRow))
    result_busv5 = np.zeros((NrRow))
    result_busv6 = np.zeros((NrRow))
    result_busv7 = np.zeros((NrRow))
    result_busv8 = np.zeros((NrRow))
    result_busv9 = np.zeros((NrRow))
    result_busv10 = np.zeros((NrRow))
    result_busv11 = np.zeros((NrRow))
    result_busv12 = np.zeros((NrRow))
    result_busv13 = np.zeros((NrRow))
    result_busv14 = np.zeros((NrRow))
    result_busv15 = np.zeros((NrRow))
    result_busv16 = np.zeros((NrRow))
    result_busv17 = np.zeros((NrRow))
    result_busv18 = np.zeros((NrRow))
    result_busv19 = np.zeros((NrRow))
    result_busv20 = np.zeros((NrRow))
    result_busv21 = np.zeros((NrRow))
    result_busv22 = np.zeros((NrRow))
    result_busv23 = np.zeros((NrRow))
    result_busv24 = np.zeros((NrRow))
    result_busv25 = np.zeros((NrRow))
    result_busv26 = np.zeros((NrRow))
    result_busv27 = np.zeros((NrRow))
    result_busv28 = np.zeros((NrRow))
    result_busv29= np.zeros((NrRow))
    result_busv30 = np.zeros((NrRow))
    result_busv31 = np.zeros((NrRow))
    result_busv32= np.zeros((NrRow))
    result_busv33 = np.zeros((NrRow))
    result_busv34 = np.zeros((NrRow))
    result_busv35 = np.zeros((NrRow))
    result_busv36 = np.zeros((NrRow))
    result_busv37 = np.zeros((NrRow))
    result_busv38 = np.zeros((NrRow))
    result_busv39 = np.zeros((NrRow))

    result_busva1 = np.zeros((NrRow))
    result_busva2 = np.zeros((NrRow))
    result_busva3 = np.zeros((NrRow))
    result_busva4 = np.zeros((NrRow))
    result_busva5 = np.zeros((NrRow))


    #get results for each time step
    for i in range(NrRow):
        result_time[i] = app.ResGetData(elmRes,i,ColIndex_time)[1]
        result_name[i] = loc_fl.loc_name
        result_scalle[i]= np.around(c,4)
        result_ct[i]=sec
        result_loca[i]=loca
        #location[i] = ColIndex_loca
        result_angle1[i] = app.ResGetData(elmRes,i,ColIndex_angle1)[1]
        result_angle2[i] = app.ResGetData(elmRes,i,ColIndex_angle2)[1]
        result_angle3[i] = app.ResGetData(elmRes,i,ColIndex_angle3)[1]
        result_angle4[i] = app.ResGetData(elmRes,i,ColIndex_angle4)[1]
        result_angle5[i] = app.ResGetData(elmRes,i,ColIndex_angle5)[1]
        result_angle6[i] = app.ResGetData(elmRes,i,ColIndex_angle6)[1]
        result_angle7[i] = app.ResGetData(elmRes,i,ColIndex_angle7)[1]
        result_angle8[i] = app.ResGetData(elmRes,i,ColIndex_angle8)[1]
        result_angle9[i] = app.ResGetData(elmRes,i,ColIndex_angle9)[1]
        result_angle10[i] = app.ResGetData(elmRes,i,ColIndex_angle10)[1]
        
        result_speed1[i] = app.ResGetData(elmRes,i,ColIndex_speed1)[1]
        result_speed2[i] = app.ResGetData(elmRes,i,ColIndex_speed2)[1]
        result_speed3[i] = app.ResGetData(elmRes,i,ColIndex_speed3)[1]
        result_speed4[i] = app.ResGetData(elmRes,i,ColIndex_speed4)[1]
        result_speed5[i] = app.ResGetData(elmRes,i,ColIndex_speed5)[1]
        result_speed6[i] = app.ResGetData(elmRes,i,ColIndex_speed6)[1]
        result_speed7[i] = app.ResGetData(elmRes,i,ColIndex_speed7)[1]
        result_speed8[i] = app.ResGetData(elmRes,i,ColIndex_speed8)[1]
        result_speed9[i] = app.ResGetData(elmRes,i,ColIndex_speed9)[1]
        result_speed10[i] = app.ResGetData(elmRes,i,ColIndex_speed10)[1]

        result_busv1[i] = app.ResGetData(elmRes,i,ColIndex_u_bus1)[1]
        result_busv2[i] = app.ResGetData(elmRes,i,ColIndex_u_bus2)[1]
        result_busv3[i] = app.ResGetData(elmRes,i,ColIndex_u_bus3)[1]
        result_busv4[i] = app.ResGetData(elmRes,i,ColIndex_u_bus4)[1]
        result_busv5[i] = app.ResGetData(elmRes,i,ColIndex_u_bus5)[1]
        result_busv6[i] = app.ResGetData(elmRes,i,ColIndex_u_bus6)[1]
        result_busv7[i] = app.ResGetData(elmRes,i,ColIndex_u_bus7)[1]
        result_busv8[i] = app.ResGetData(elmRes,i,ColIndex_u_bus8)[1]
        result_busv9[i] = app.ResGetData(elmRes,i,ColIndex_u_bus9)[1]
        result_busv10[i] = app.ResGetData(elmRes,i,ColIndex_u_bus10)[1]
        result_busv11[i] = app.ResGetData(elmRes,i,ColIndex_u_bus11)[1]
        result_busv12[i] = app.ResGetData(elmRes,i,ColIndex_u_bus12)[1]
        result_busv13[i] = app.ResGetData(elmRes,i,ColIndex_u_bus13)[1]
        result_busv14[i] = app.ResGetData(elmRes,i,ColIndex_u_bus14)[1]
        result_busv15[i] = app.ResGetData(elmRes,i,ColIndex_u_bus15)[1]
        result_busv16[i] = app.ResGetData(elmRes,i,ColIndex_u_bus16)[1]
        result_busv17[i] = app.ResGetData(elmRes,i,ColIndex_u_bus17)[1]
        result_busv18[i] = app.ResGetData(elmRes,i,ColIndex_u_bus18)[1]
        result_busv19[i] = app.ResGetData(elmRes,i,ColIndex_u_bus19)[1]
        result_busv20[i] = app.ResGetData(elmRes,i,ColIndex_u_bus20)[1]
        result_busv21[i] = app.ResGetData(elmRes,i,ColIndex_u_bus21)[1]
        result_busv22[i] = app.ResGetData(elmRes,i,ColIndex_u_bus22)[1]
        result_busv23[i] = app.ResGetData(elmRes,i,ColIndex_u_bus23)[1]    
        result_busv24[i] = app.ResGetData(elmRes,i,ColIndex_u_bus24)[1]
        result_busv25[i] = app.ResGetData(elmRes,i,ColIndex_u_bus25)[1]
        result_busv26[i] = app.ResGetData(elmRes,i,ColIndex_u_bus26)[1]
        result_busv27[i] = app.ResGetData(elmRes,i,ColIndex_u_bus27)[1]
        result_busv28[i] = app.ResGetData(elmRes,i,ColIndex_u_bus28)[1]
        result_busv29[i] = app.ResGetData(elmRes,i,ColIndex_u_bus29)[1]
        result_busv30[i] = app.ResGetData(elmRes,i,ColIndex_u_bus30)[1]
        result_busv31[i] = app.ResGetData(elmRes,i,ColIndex_u_bus31)[1]
        result_busv32[i] = app.ResGetData(elmRes,i,ColIndex_u_bus32)[1]
        result_busv33[i] = app.ResGetData(elmRes,i,ColIndex_u_bus33)[1]
        result_busv34[i] = app.ResGetData(elmRes,i,ColIndex_u_bus34)[1]
        result_busv35[i] = app.ResGetData(elmRes,i,ColIndex_u_bus35)[1]
        result_busv36[i] = app.ResGetData(elmRes,i,ColIndex_u_bus36)[1]
        result_busv37[i] = app.ResGetData(elmRes,i,ColIndex_u_bus37)[1]
        result_busv38[i] = app.ResGetData(elmRes,i,ColIndex_u_bus38)[1]
        result_busv39[i] = app.ResGetData(elmRes,i,ColIndex_u_bus39)[1]

        result_busva1[i] = app.ResGetData(elmRes,i,ColIndex_u_busa1)[1]
        result_busva2[i] = app.ResGetData(elmRes,i,ColIndex_u_busa2)[1]
        result_busva3[i] = app.ResGetData(elmRes,i,ColIndex_u_busa3)[1]
        result_busva4[i] = app.ResGetData(elmRes,i,ColIndex_u_busa4)[1]
        result_busva5[i] = app.ResGetData(elmRes,i,ColIndex_u_busa5)[1]
        
    
    results = pd.DataFrame()
    time=pd.DataFrame()
    time['time'] = result_time
    
    name=pd.DataFrame()
    name['name'] = result_name
    name['loadscalling'] = result_scalle
    name['ct'] = result_ct
    name['loca'] = result_loca
    
    
    results['G01']=result_angle1
    results['G02']=result_angle2
    results['G03']=result_angle3
    results['G04']=result_angle4
    results['G05']=result_angle5
    results['G06']=result_angle6
    results['G07']=result_angle7
    results['G08']=result_angle8
    results['G09']=result_angle9
    results['G10']=result_angle10
    results['dfrot']=results.apply(lambda x: x.max() - x.min(), axis = 1)
    results['stability']=results['dfrot'].apply(lambda x: 1 if x>180 else 0 )
    
    speed=pd.DataFrame()
    speed['G01s']=result_speed1
    speed['G02s']=result_speed2
    speed['G03s']=result_speed3
    speed['G04s']=result_speed4
    speed['G05s']=result_speed5
    speed['G06s']=result_speed6
    speed['G07s']=result_speed7
    speed['G08s']=result_speed8
    speed['G09s']=result_speed9
    speed['G10s']=result_speed10

    busv=pd.DataFrame()
    busv['busv1']=result_busv1
    busv['busv2']=result_busv2
    busv['busv3']=result_busv3
    busv['busv4']=result_busv4
    busv['busv5']=result_busv5
    busv['busv6']=result_busv6
    busv['busv7']=result_busv7
    busv['busv8']=result_busv8
    busv['busv9']=result_busv9
    busv['busv10']=result_busv10
    busv['busv11']=result_busv11
    busv['busv12']=result_busv12
    busv['busv13']=result_busv13
    busv['busv14']=result_busv14
    busv['busv15']=result_busv15
    busv['busv16']=result_busv16
    busv['busv17']=result_busv17
    busv['busv18']=result_busv18
    busv['busv19']=result_busv19
    busv['busv20']=result_busv20
    busv['busv21']=result_busv21
    busv['busv22']=result_busv22
    busv['busv23']=result_busv23
    busv['busv24']=result_busv24
    busv['busv25']=result_busv25
    busv['busv26']=result_busv26
    busv['busv27']=result_busv27
    busv['busv28']=result_busv28
    busv['busv29']=result_busv29
    busv['busv30']=result_busv30
    busv['busv31']=result_busv31
    busv['busv32']=result_busv32
    busv['busv33']=result_busv33
    busv['busv34']=result_busv34
    busv['busv35']=result_busv35
    busv['busv36']=result_busv36
    busv['busv37']=result_busv37
    busv['busv38']=result_busv38
    busv['busv39']=result_busv39

    busv['busva1']=result_busva1
    busv['busva2']=result_busva2
    busv['busva3']=result_busva3
    busv['busva4']=result_busva4
    busv['busva5']=result_busva5
    
    time=pd.concat([time,name],axis=1,ignore_index=False)
    results=pd.concat([time,results],axis=1,ignore_index=False)
    results=pd.concat([results,busv],axis=1,ignore_index=False)
    results=pd.concat([results,speed],axis=1,ignore_index=False)
    return results

  
def runbus(obj):
  buses = app.GetCalcRelevantObjects('*.ElmTerm')
  lines = app.GetCalcRelevantObjects('*.ElmLne')
  CT=[0.18,0.19,0.20,0.21,0.22,0.23]
  scale=[0.6,0.65,0.7,0.75,0.80,0.85,0.9,0.95,1,1.05,1.1,1.15,1.2]
  random_element = obj
  element=random_element.loc_name
  for bus in buses:
    if bus == random_element:
      loc_fl= bus
      for sec in CT:
        time = 0
        clearTime = sec
        faultType = 0
        faulClear = 4
        for loadscall in scale:
          loadscalling(loadscall)
          setupSimultation(comInc, comSim,0.1,5,0.01)
          addShcEvent(loc_fl, time, faultType)
          addShcEvent(loc_fl, clearTime, faulClear)
          runSimulation(comInc, comSim)
          location=50
                
          RES = getResults(obj,sec,location)
          app.PrintPlain(RES)
          Data=pd.DataFrame(RES)
          app.PrintPlain(list(Data))
          filename = 'main111.csv'
          Data.to_csv(filename, mode='a', header=True )  

def runline(obj):
  buses = app.GetCalcRelevantObjects('*.ElmTerm')
  lines = app.GetCalcRelevantObjects('*.ElmLne')
  loca=[5,25,50,75,95]
  CT=[0.18,0.19,0.20,0.21,0.22,0.23]
  scale=[0.6,0.65,0.7,0.75,0.80,0.85,0.9,0.95,1,1.05,1.1,1.15,1.2]
  random_element = obj
  element=random_element.loc_name
  for line in lines:
    if line == random_element:
      loc_fl= line
      for location in loca:
        loc_fl.SetAttribute('e:fshcloc',location)
        for sec in CT:
          time = 0
          clearTime = sec
          faultType = 0
          faulClear = 4
          for loadscall in scale:
            loadscalling(loadscall)
            setupSimultation(comInc, comSim,0.1,5,0.01)
            addShcEvent(loc_fl, time, faultType)
            addShcEvent(loc_fl, clearTime, faulClear)
            runSimulation(comInc, comSim)
                
            RES = getResults(obj,sec,location)
            app.PrintPlain(RES)
            Data=pd.DataFrame(RES)
            app.PrintPlain(list(Data))
            filename = 'line.csv'
            Data.to_csv(filename, mode='a', header=True )
            clearSimEvents()
  
      

#################################################################################  
projName='39 Bus New England System2'
study_case='Simulation Fault Bus 31 UnStable.IntCase'

project = app.ActivateProject(projName)
proj = app.GetActiveProject()
Folder_studycase = app.GetProjectFolder('study')

##
Case = Folder_studycase.GetContents(study_case)
app.PrintPlain(Case)
comInc = app.GetFromStudyCase('ComInc')
comSim = app.GetFromStudyCase('ComSim')

lines = app.GetCalcRelevantObjects("*.ElmLne")
buses = app.GetCalcRelevantObjects('*.ElmTerm')
Full_tit = buses+lines


runbus(buses[38])
#runline(lines[0])











