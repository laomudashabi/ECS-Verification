#!/usr/bin/python
# coding=utf-8
""" This module contains the feature that executes test automation through
CD & AD API accordig to a released TestSchema.
Note that this class is in an early stage of development and will probably
change over time to make it more practical to use.

Author: Zhetong Mo
zhetong.mo@volvo.com
Updates by: zhetong.mo@volvo.com
            
"""

from win32com.client import Dispatch
import win32api
import xml.etree.ElementTree as ET
import logging
import sys
import os
import threading


ShortName={'DriveLevelControl':['DLC_TC','DriveLevelControl',1],'ParameterCheck':['PRC_TC','ParameterCheck',1],'ProhibitControl':['PC_TC','ProhibitControl',1],
           'DowngradeMode':['DGM_TC','DowngradeMode',1],'AirDumpFunction':['ADF_TC','AirDumpFunction',1],'ECSStandby':['STB_TC','ECSStandby',1],
           'FerryFunction':['FF_TC','FerryFunction',1],'Kneeling':['KNL_TC','Kneeling',1],'LoadingLevelControl':['LEC_TC','LoadingLevelControl',1]
           ,'UnevenLoadHandling':['ULH_TC','UnevenLoadHandling',1]}

def CreateLogger(loggerFilePath):
    global logger
    try:
        logger = logging.getLogger('ExecTestScope')
        if not len(logger.handlers):
            logger.setLevel(logging.DEBUG)
            # create file handler which logs even debug messages
            fh = logging.FileHandler(loggerFilePath)
            fh.setLevel(logging.DEBUG)
            # create console handler with a higher log level
            ch = logging.StreamHandler(sys.stdout)
            ch.setLevel(logging.DEBUG)
            # create formatter and add it to the handlers
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            fh.setFormatter(formatter)
            ch.setFormatter(formatter)
            # add the handlers to the logger
            logger.addHandler(fh)
            logger.addHandler(ch)
            logger.debug("Logger created")
            return 0
    except Exception as e:
        print("Error: Logger could not be created: "+str(sys.exc_value))
        logger.error(str(e))
        return 1

def ExitLogger():
    global logger
    handlers = list(logger.handlers)
    for handler in handlers:
        handler.flush()
        handler.close()
        logger = None 

# To extract all necessary information from TestSchema.xml
# Input: TestSchema.xml
# Output: Result(0=success, 1=open xml failed, 2=extract info failed), TestCases(dictionary)
def TestScope(TestSchema):
    result = 0
    doc = None
    try:
        f = open(TestSchema, 'r')
    except:
        logger.error("Error opening XML-file")
        result = 1
        return (result, None)
    try:
        xmlData = f.read()
        f.close()
        root = ET.fromstring(xmlData)
        Requirements = root.find('RequirementGroups')
        TestCases = {}
        for Requirement in Requirements.getchildren():  
            TCName = str(Requirement.attrib['key'])
            for Baseline in Requirement.find('SoftwareBaselines').getchildren():
                BLName = str(Baseline.attrib['key'])
                if BLName not in TestCases:
                    TestCases.update({BLName:{TCName:{}}})
                else:
                    TestCases[BLName].update({TCName:{}})
                Vdic = {}
                for Vehicle in Baseline.find('VehicleVariants').getchildren():
                    VName = str(Vehicle.attrib['key'])
                    TClist = []
                    for TC in Vehicle.find('RequiredTCs').findall('TC'):
                        TClist.append(TC.text)
                    Vdic.update({VName:TClist})
                TestCases[BLName][TCName].update(Vdic)                                    
    except EX as e:
        if f:
            f.close()
        logger.error(str(e))
        logger.error("Error parsing XML code from XML-file")
        result = 2        
        return (result, None)
    xmlData = None
    root = None
    Requirements = None
    return (result, TestCases)

#To create executation plan based on different trucks
def CreateExecPlan(TestCases):
    result = 0
    ExecPlan = {}
    try:
        for BL in TestCases:
            ExecPlan.update({BL:{}})
            for Req in TestCases[BL]:
                for truck in TestCases[BL][Req]:
                    if truck not in ExecPlan[BL]:
                        ExecPlan[BL].update({truck:{Req:TestCases[BL][Req][truck]}})
                    else:
                        ExecPlan[BL][truck].update({Req:TestCases[BL][Req][truck]})              
    except Exception as e:
        result = 1
        logger.error(str(e))
        logger.error("Error converting test scope into truck based executation plan.")
        raise
    return (result, ExecPlan)

def ConfigAD(ProjName, CurrentVec):
    try:
        # Create the COM Server
        AudObj = Dispatch("AutomationDesk.TAM.5.6")
        # Show the user interface of AutomationDesk
        AudObj.Visible = True
        # Get the Projects collection
        ProjsObj = AudObj.Projects
        ProjObj = None
        for Name in ProjsObj.Names:
            if Name in ProjName:
                ProjObj = ProjsObj.Item(Name)
        if ProjObj == None:
            ProjObj = ProjsObj.Load(ProjName)
        # Get or create the right folder as object
        FolderObj = ProjObj.SubBlocks.Item("TestCase")
        # Get the ET settings
        ETObj = ProjObj.DataObjects.Item("ConsoleET")
        console = ETObj.ChildDataObjects.Item('PythonLibPath').Value
        etexe = ETObj.ChildDataObjects.Item('ConsoleETBinary').Value
        ETsetting = ETObj.ChildDataObjects.Item('ConsoleETXml').Value
        # Set CurrentVehicle
        DataObj = ProjObj.DataObjects.Item("GlobalSettings")
        CV = DataObj.ChildDataObjects.Item('CurrentVehicle')
        CV.Value = CurrentVec
        return (0,console,etexe,ETsetting,ProjObj)
    except Exception as e:
        logger.error(str(e))
        logger.error("Error accessing AutomationDesk project " + ProjName)
        return (1,None,None,None,None)
        raise

def ConfigTP_CD(ProjObj, ExecPlan, TestObject, logger):
    try: # Config AutomationDesk project
        # Parce baseline from project name
        BL = ''
        for BLName in ExecPlan:
            if BLName.replace('.','_').replace('/','-') in ProjObj.Name:
                BL = BLName
        if BL != '':
            # Enable all the TCs which are not picked
            FolderObj = ProjObj.SubBlocks.Item("TestCase")
            for Req in ExecPlan[BL][TestObject]:
                if "Manual" not in Req and ExecPlan[BL][TestObject][Req][0] != 'all':
                    Enabled = []
                    for i in ExecPlan[BL][TestObject][Req]:
                        j = ShortName[Req][0] + i[-3:]
                        Enabled.append(j)
                    FolderObj1 = FolderObj.SubBlocks.Item(Req)
                    NoOfElements = FolderObj1.SubBlocks.Count
                    TC = [0]*NoOfElements
                    for i in range(0,NoOfElements):
                        TC[i] = FolderObj1.SubBlocks.Item(i)
                        if TC[i].Name in Enabled:
                            TC[i].IsEnabled = 1
            ProjObj.Save()
            logger.info("AutomationDesk project " + ProjObj.Name + " configured.")
        else:
            logger.error("Error parcing baseline from the AD project " + ProjObj.Name)
    except Exception as e:
        logger.error(str(e))
        logger.error("Error configuring test project according to ExecPlan.")
        raise

    try: # To configure ControlDesk with the current test object
        application =  Dispatch("ControlDeskNG.Application") # Open a COM connection to ControlDesk.
        platform = application.ActiveExperiment.Platforms[0] # Get the sole platform
        datasets = platform.ActiveVariableDescription.DataSets # Get datasets under the platform
        
        if application.CalibrationManagement.StartOnlineCalibration() == 0:
            for i in range(0,datasets.Count):
                if TestObject in str(datasets[i].FileName):
                    datasets.Item(i).WriteToHardware()
                    datasets.Item(i).WriteToHardware() # Force the picked test object's dataset to be loaded
                else:
                    datasets.Item(i).Close()
            logger.info("ControlDesk project configured.")
        else:
            logger.error("ControlDesk cannot go online.")
    except:
        logger.error("Error accessing ControlDesk.")
        raise
    finally:
        application=None;platform=None;datasets=None

def ExecTP(ProjObj):   
    #Execute the entire project
    try:
        print "====================== Start executing project =================="
        ProjObj.Execute(BLName)
    except:
        logger.error("Error executing AD project.")
        raise
    finally:
        ProjObj=None

def DisableTC(ProjObj): # Disable all the TCs to prepare for next test object
    try:
        FolderObj = ProjObj.SubBlocks.Item("TestCase")
        for i in range(0,FolderObj.SubBlocks.Count):
            FolderObj1 = FolderObj.SubBlocks.Item(i)
            NoOfElements = FolderObj1.SubBlocks.Count
            for j in range(0,NoOfElements):
                TC = FolderObj1.SubBlocks.Item(j)
                TC.IsEnabled = 0
        ProjObj.Save()
    except:
        logger.error("Error disabling all test cases.")
        raise
    finally:
        ProjObj=None

def DownloadSW(TestObject,SWPath,Node,etexe,ETsetting,logger):
    files = []
    for file in os.listdir(SWPath): # Find the SW package matching test object
        if file.endswith('.zip') and TestObject in file:
            files.append(os.path.join(SWPath, file))
    if len(files) == 1:
        swlist = [(Node, files[0])]
        try:
            ET = ETConsole.ConsoleET(etexe,ETsetting)
            if ET.start() == 0:
                if ET.login() == 0:
                    if ET.connect() == 0:
                        if ET.listEcu()[0] == 0:
                            status = ET.downloadSoftware(swlist)     
                            if status == 0:
                                logger.info(str(files[0])+" sucessfully donloaded to node(s) "+Node)
                                ET.vehicleReset(True)
                                if ET.writeParam({'P1D62': ['1'], 'P1D63': ['1']})[0] == 0:
                                    logger.info("Debug data has been enabled.")
                                else:
                                    logger.warning("Enable debug data failed.")
                            else:
                                logger.error("Software download failed. Return code: "+str(status))
            else:
                logger.error("Error: Could not start Console ET.")
        except Exception as e:
            logger.error(str(e))
            raise
        finally:
            ET.exit()
    elif len(files) == 0:
        logger.error("No SW package found for "+TestObject+" in path: "+SWPath+".")
    else:
        logger.error("Multiple SW packages found for "+TestObject+" in path: "+SWPath+".")

def main(TestSchema,ProjName,TestObject,SWPath,Node):
    global AudObj,ProjObj,ETConsole,cwd,console,etexe,ETsetting
    try:
        cwd = os.getcwd()
        CreateLogger(cwd+'\\Exec.log')
        (result,TestCases) = TestScope(TestSchema)
        if result == 0:
            (result,ExecPlan) = CreateExecPlan(TestCases)  
        else:
             logger.error("Error parcing " + TestSchema)
    except Exception as e:
        logger.error(str(e))
        raise
    (status,console,etexe,ETsetting,ProjObj) = ConfigAD(ProjName, TestObject)
    if result == 0 and status == 0:
        sys.path.insert(1,console)
        import console_et as ETConsole
        try:
            newThread = threading.Thread(target=DownloadSW, args =(TestObject,SWPath,Node,etexe,ETsetting,logger)) # Create a new thread to download SW.
            newThread.start() # Start the thread
            ConfigTP_CD(ProjObj,ExecPlan,TestObject,logger) # Configure CD & AD projects in the current thread
        except Exception as e:
            logger.error(str(e))
            raise
        finally:
            if ProjObj != None:
                DisableTC(ProjObj)
            AudObj=None
            ExitLogger()
    else:
        logger.error("Error configuring AutomationDesk.")
    
main('C:\\Users\\a269028\\Desktop\\TestScope.xml',"C:\\ECS_HIL_Simulator\\ECS\\HIL\\AutomationDesk\\Projects\\T2_B_T1_A_TestExecutionProject\\T2_A_24_w2109_17V1.adp",
"FH-2267","C:\\ECS_HIL_Simulator\\ECS\\HIL\\ECU_SW\\T2_B\\",'32,34')

#Only run if called from cmd
if __name__ == '__main__':
    main(sys.argv[1],sys.argv[2],sys.argv[3],sys.argv[4],sys.argv[5])