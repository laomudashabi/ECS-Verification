""" This module contains the feature that assembles AD projects through
AD API accordig to a released TestSchema.
Note that this class is in an early stage of development and will probably
change over time to make it more practical to use.

Author: Zhetong Mo
zhetong.mo@volvo.com
Updates by: zhetong.mo@volvo.com
            
"""
import win32com.client
import win32api
import xml.etree.ElementTree as ET
import logging
import sys
import os

ShortName={'DriveLevelControl':['DLC_TC','DriveLevelControl',1],'ParameterCheck':['PRC_TC','ParameterCheck',1],'ProhibitControl':['PC_TC','ProhibitControl',1],
           'DowngradeMode':['DGM_TC','DowngradeMode',1],'AirDumpFunction':['ADF_TC','AirDumpFunction',1],'ECSStandby':['STB_TC','ECSStandby',1],
           'FerryFunction':['FF_TC','FerryFunction',1],'Kneeling':['KNL_TC','Kneeling',1],'LoadingLevelControl':['LEC_TC','LoadingLevelControl',1]
           ,'UnevenLoadHandling':['ULH_TC','UnevenLoadHandling',1]}

def CreateLogger(loggerFilePath):
    try:
        global logger
        logger = logging.getLogger('TestScope')
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
        print "Error: Logger could not be created: "+str(sys.exc_value)
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
    return result, TestCases

# To create the test project according to the test scope.
# Return a list that contains all created project paths
def CreateTP(TestCases,TemplatePath,ProjectPath):    
    AudObj=None;ProjsObj=None;ProjObj=None; NewProjects = []
    for BLName in TestCases:
        # Check if baseline name is supported
        if 'T1_' in BLName or 'T2_' in BLName:
            NewProject = ProjectPath.replace('\"', "") + '\\' + BLName.replace('.','_').replace('/','-') +'.adp'
            logger.info("Creating project for baseline " + BLName.replace('.','_').replace('/','-'))
        else:
            logger.error("Baseline name " + BLName.replace('.','_').replace('/','-') + " not supported.")
            exit()
        NewProjects.append(NewProject)
        try:
            # Create the COM Server
            AudObj = win32com.client.Dispatch("AutomationDesk.TAM.5.6")
            # Show the user interface of AutomationDesk
            AudObj.Visible = True
            # Get the Projects collection
            ProjsObj = AudObj.Projects
            ProjName = TemplatePath
            for Name in ProjsObj.Names:
                if Name in ProjName:
                    print "====================== Now closing exisiting project "+Name+" =================="
                    PeojObj = ProjsObj.Item(Name)
                    PeojObj.Close()
            ProjObj = ProjsObj.ImportProject(ProjName,0,1)
            # Get or create the right folder as object
            FolderObj = ProjObj.SubBlocks.Item("TestCase")
            # Get the Libraries collection
            LibsObj = AudObj.Libraries
        except:
            logger.error("Error opening template project in AD.")
            raise
        try:
            for FuncName in TestCases[BLName]: 
                if FuncName in FolderObj.SubBlocks.Names:
                    FolderObj1 = FolderObj.SubBlocks.Item(FuncName)
                elif "Manual" not in FuncName:
                    StdLibObj = LibsObj.Item("Standard")
                    FolderTemplObj = StdLibObj.SubBlocks.Item("Folder")
                    FolderObj1 = FolderObj.SubBlocks.Create(FolderTemplObj)
                    FolderObj1.Name = FuncName
                # Get the Custom Library
                if FuncName in ShortName:
                    logger.info("Copying test cases for Function " + FuncName)
                    CustomLibObj = LibsObj.Item(ShortName[FuncName][1])
                    NoOfElements = CustomLibObj.SubBlocks.Count
                    print ShortName[FuncName][1]+" Custom Library contains %i element(s)"\
                                        %NoOfElements
                    TC = [0]*NoOfElements
                    for i in range (0,NoOfElements):
                        TC[i] = CustomLibObj.SubBlocks.Item(i)                    
                    for i in range (0,NoOfElements):
                        if TC[i].Type == 2:
                            SequenceObj = FolderObj1.SubBlocks.Create(TC[i])
                            SequenceObj.Name = TC[i].Name
                            SequenceObj.IsCollapsed = True
                            SequenceObj.IsEnabled = 0
                    FolderObj1.IsCollapsed = True
                    logger.info(FuncName + " test cases collected.")
        except:
            logger.error("Error linking test cases from library:"+str(FuncName))
            raise
        # Save and close the project           
        try:
            print "====================== Now saving new project as "+NewProject+" =================="
            for Name in ProjsObj.Names:
                if Name in NewProject:
                    PeojObj = ProjsObj.Item(Name)
                    PeojObj.Close()
            ProjObj.SaveAs(NewProject,1)
            logger.info(BLName.replace('.','_').replace('/','-') + " test project generated.")
        except:
            logger.error("Error saving new project in AD.")
            raise
    logger.info("Test projects sucessfully generated according to the TestScope.")
    return NewProjects
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
    return result, ExecPlan

# Disable all TCs which are not in the ExecPlan and execute the project
def ExecTP(ProjName, ExecPlan, CurrentVec):
    try:
        # Create the COM Server
        AudObj = win32com.client.Dispatch("AutomationDesk.TAM.5.6")
        # Show the user interface of AutomationDesk
        AudObj.Visible = True
        # Get the Projects collection
        ProjsObj = AudObj.Projects
        for Name in ProjsObj.Names:
            if Name in ProjName:
                ProjObj = ProjsObj.Item(Name)
        if ProjObj == None:
            ProjObj = ProjsObj.Load(ProjName)
        # Get or create the right folder as object
        FolderObj = ProjObj.SubBlocks.Item("TestCase")
        # Set CurrentVehicle
        DataObj = ProjObj.DataObjects.Item("GlobalSettings")
        CV = DataObj.ChildDataObjects.Item('CurrentVehicle')
        CV.Value = CurrentVec
        # Parce baseline from project name
        BL = ''
        for BLName in ExecPlan:
            if BLName.replace('.','_').replace('/','-') in ProjName:
                BL = BLName
        if BL != '':
            # Disable all the TCs which are not picked
            for Req in ExecPlan[BL][CurrentVec]:
                if "Manual" not in Req and ExecPlan[BL][CurrentVec][Req][0] != 'all':
                    Enabled = []
                    for i in ExecPlan[BL][CurrentVec][Req]:
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
            print "====================== Baseline name: " + BL + " =================="
            
##            #Execute the entire project
##            try:
##                print "====================== Start executing project =================="
##                ProjObj.Execute(BLName)
##            except:
##                logger.error("Error executing AD project.")
##                raise
        else:
            logger.error("Error parcing baseline from the AD project" + ProjName)
    except Exception as e:
        logger.error(str(e))
        logger.error("Error configuring test project according to ExecPlan.")
        raise

def ExitAD():
    if ProjObj is not None:
        ProjObj.Close()
    if AudObj is not None:
        AudObj.Quit()
    ProjObj = None
    ProjsObj = None
    win32api.Sleep(5000)
    AudObj = None
    print "====================== Project assembling successfully finished =================="

def main(TestSchema,TemplatePath,ProjectPath):
    try:
        cwd = os.getcwd()
        CreateLogger(cwd+'\\ADAPI.log')
        (result,TestCases) = TestScope(TestSchema)
        (result,ExecPlan) = CreateExecPlan(TestCases)    
        NewProjects = CreateTP(TestCases,TemplatePath,ProjectPath)
        for NewProject in NewProjects:
            for Baseline in ExecPlan:
                if Baseline.replace('.','_').replace('/','-') in NewProject and ExecPlan[Baseline] != {}:
                    Vehicle = list(ExecPlan[Baseline].keys())[0]
                    ExecTP(NewProject, ExecPlan, Vehicle)
                    break
            else:
                continue
    except Exception as e:
        logger.error(str(e))
        raise
    finally:
        ExitLogger()

       
##TestSchema = 'C:\\Users\\a269028\\Desktop\\TestScope.xml'
##ProjectPath = 'C:\\Users\\a269028\\Desktop\\'
##TemplatePath = 'C:\\Users\\a269028\\Desktop\\SE-Tool_Integration_APP\\Template_TestExecutionProject_AD5-6.zip'
##main(TestSchema,TemplatePath,ProjectPath)

#Only run if called from cmd
if __name__ == '__main__':
    main(sys.argv[1],sys.argv[2],sys.argv[3])


