import subprocess   # For launching the analysis engine
import win32com.client
from win32com.client import gencache

# win32com.client.constants.* come from the makepy type-library cache, so generate it for the two
# LARSA libraries this example uses. Then force late binding so the objects keep their real
# method/property names when reading data from LARSA (an early-bound cache hides element fields 
# behind base COM interfaces).
gencache.EnsureDispatch("LarsaData.clsUnits")
gencache.EnsureDispatch("LarsaElements.clsProject")
gencache.EnsureDispatch("LarsaElements.clsAnalysisResults")
gencache.GetClassForCLSID = lambda clsid: None  

def CreateObject(class_name):
    return win32com.client.Dispatch(class_name)

def columnHeaders(analysis, dataType):
    # Results spreadsheet headers (with units).
    try:
        info = win32com.client.Record("RESULTS_SPREAD_INFO", analysis.ResultGroups)
        info.rq.dataType = dataType
        info.display = 0
        return [h.replace("\r\n", " ") for h in analysis.getSpreadsheetColumnHeaders(info, [])[1] if isinstance(h, str)]
    except Exception:
        # pywin32 312 has an issue where writing any field of a COM struct raises TypeError 
        # (pywin32 <= 311 returns the labelled headers as expected).
        return []

# ----- Generate the model --------------------------------
project = CreateObject("LarsaElements.clsProject")

# Set units for the 9 units categories:
#   Input Data -> Units: 0 = MATERIAL_UNITS,  1 = Section_UNITS,  2 = Spring_units,  
#       3 = COORDINATE_UNITS, 4 = LOAD_UNITS,  5 = mass_units;
#   Results -> Units: 6 = DISPLACEMENT_UNITS,  7 = FORCE_UNITS,  8 = STRESS_UNITS.
for category in range(9):
    project.units.SetUnits(category, True, 
        project.units.UnitFromString(win32com.client.constants.UNIT_LENGTH, "in"), 
        project.units.UnitFromString(win32com.client.constants.UNIT_FORCE, "kip"), 
        project.units.UnitFromString(win32com.client.constants.UNIT_TEMPERATURE, "F"), 
        project.units.UnitFromString(win32com.client.constants.UNIT_ANGLE, "rad"))

material = CreateObject("LarsaElements.clsMaterial")
material.Name = "Steel"
material.modulusOfElasticity = 29000.0    
material.unitWeight = 0.000284
project.materials.append(material)

section = CreateObject("LarsaElements.clsSection")   
section.Name = "Section 1"
section.sectionArea = 10.0
section.inertiaY = 100.0
section.inertiaZ = 100.0
section.torsionj = 50.0
project.sections.append(section)

joint1 = CreateObject("LarsaElements.clsJoint")
joint1.location.setCoordinates(0, 0, 0)
project.joints.append(joint1)
joint1.constraint = "111111"   # All DOFs are fixed: tx ty tz rx ry rz (1 = fixed, 0 = free)

joint2 = CreateObject("LarsaElements.clsJoint")
joint2.location.setCoordinates(120, 0, 0)
project.joints.append(joint2)

joint3 = CreateObject("LarsaElements.clsJoint")
joint3.location.setCoordinates(240, 0, 0)
project.joints.append(joint3)

member1 = CreateObject("LarsaElements.clsMember")
member1.Setjoint(1, joint1)
member1.Setjoint(2, joint2)
member1.material = material
member1.sectionID(1, section.number)    # 1 = Section as Start, 2 = Section at End (if different)
project.members.append(member1)

member2 = CreateObject("LarsaElements.clsMember")
member2.Setjoint(1, joint2)
member2.Setjoint(2, joint3)
member2.material = material
member2.sectionID(1, section.number)
project.members.append(member2)

lc1 = CreateObject("LarsaElements.clsLoadCase")
lc1.Name = "Self weight"
project.primaryLoadCases.append(lc1)
lc1.weightFactor.z = -1

lc2 = CreateObject("LarsaElements.clsLoadCase")
lc2.Name = "Uniform load"
project.primaryLoadCases.append(lc2)
for i in range(1, project.members.Count + 1):
    m = project.members.itemByIndex(i)
    mload = CreateObject("LarsaElements.clsMemberLoad")
    lc2.MemberLoads.append(mload)
    mload.member = m
    mload.loadType = win32com.client.constants.UniformForce
    mload.loadDir = win32com.client.constants.Global_z
    mload.startW = -0.1

group = CreateObject("LarsaElements.clsGeoGroup")
group.Name = "Whole structure"
project.geoGroups.append(group)
group.addObject(win32com.client.constants.GGA_MEMBERS, member1.number)  
group.addObject(win32com.client.constants.GGA_MEMBERS, member2.number)   

stage = CreateObject("LarsaElements.clsStage")
project.stages.append(stage)         
stage.Name = "Stage 1"

step1 = CreateObject("LarsaElements.clsStageStep")
stage.steps.append(step1)
step1.Name = "Construction"
step1.stageType = win32com.client.constants.ST_CONSTRUCTION
step1.constructionMethod = win32com.client.constants.SCM_STANDARD # Joint location initialization
step1.analysisMethod = win32com.client.constants.AT_NONLINEAR_STATIC
step1.groupAdd(group)
step1.caseAdd(lc1, 1.0)

step2 = CreateObject("LarsaElements.clsStageStep")
stage.steps.append(step2)
step2.Name = "Loading"
step2.stageType = win32com.client.constants.ST_CONSTRUCTION
step2.constructionMethod = win32com.client.constants.SCM_STANDARD
step2.analysisMethod = win32com.client.constants.AT_NONLINEAR_STATIC
step2.caseAdd(lc2, 1.0)

project.analysisType = win32com.client.constants.AT_STAGE_STANDARD   # Staged construction
project.analysisStageStart = 1
project.analysisStageEnd = project.stages.Count

filename = r"C:\Users\LARSA\Documents\LARSA Projects\test.lar"
lar_format_version = (8, 9, 0)
project.SaveToFile(filename, win32com.client.constants.FILE_FORMAT_LAR6A,
                   *lar_format_version, None)

# ----- Run the analysis (standalone engine, no GUI) ------
engine_exe = r"C:\Program Files (x86)\Larsa 2000\LarsaEngine64.exe"
subprocess.run([engine_exe, "/run", filename], check=True)

# ----- Read the results ----------------------------------
analysis = CreateObject("LarsaElements.clsAnalysisResults")
analysis.Load(project)
project.ReadFromFile(filename, analysis, False)
analysis.Load(project)

caseName = "Stage 1: Loading"

def getResultCaseByName(caseName):
    resultCases = analysis.GetAllCases(None)
    return [c for c in resultCases if c.Name == caseName][0]
   
def getResults(resultCase, dataType, envelopeCol, envelopeAbs, incremental, 
                           inUCS, loadClass, indexes):
    data = analysis.loadData3(resultCase, dataType, envelopeCol, envelopeAbs, incremental, 
                           inUCS, loadClass, [], [], indexes)[8]
    data = analysis.convertData(dataType, [0] + list(data), project.units)[1][1:]   # convert into Results units
    headers = columnHeaders(analysis, dataType)  
    return dict(zip(headers, data)) if headers else list(data)
   
resultCase = getResultCaseByName(caseName)

envelopeCol = 0
envelopeAbs = False
incremental = False
inUCS = False
loadClass = 0

# Joint displacements
jointID = 3
dataType = win32com.client.constants.RESULTDATA_JOINT_DISPLACEMENTS
indexes = [0] * 4
indexes[1] = analysis.getJointIndexFromID(jointID)[0]
displ = getResults(resultCase, dataType, envelopeCol, envelopeAbs, incremental, 
                           inUCS, loadClass, indexes)
print("Displacements, Joint 3:", displ)

# Member sectional forces
memberID = 2
station = 0
numSegments = 1
dataType = win32com.client.constants.RESULTDATA_MEMBER_SECTIONAL_FORCES
indexes = [0] * 4
indexes[1] = analysis.getMemberIndexFromID(memberID)[0]
indexes[2] = station
indexes[3] = numSegments
sectionalForces = getResults(resultCase, dataType, envelopeCol, envelopeAbs, incremental, 
                           inUCS, loadClass, indexes)
print("Member sectional forces, Member 2, Station 0:", sectionalForces)

# Member end forces (local)
memberID = 1
memberEnd = 0   # 0 = Start (I-Joint), 1 = End (J-Joint)
dataType = win32com.client.constants.RESULTDATA_MEMBER_END_FORCES
indexes = [0] * 4
indexes[1] = analysis.getMemberIndexFromID(memberID)[0]
data_bothEnds = analysis.loadData3(resultCase, dataType, envelopeCol, envelopeAbs, incremental, 
                            inUCS, loadClass, [], [], indexes)[8]
data_bothEnds = analysis.convertData(dataType, [0] + list(data_bothEnds), project.units)[1][1:]
if memberEnd == 0:
    data = data_bothEnds[2:8]    # Forces at Start (I-Joint)
else:
    data = data_bothEnds[8:14]   # Forces at End (J-Joint)
headers = columnHeaders(analysis, dataType)[-6:]
endForcesLocal = dict(zip(headers, data)) if headers else list(data)
print("Member end forces (local), Member 1, Start:", endForcesLocal)
