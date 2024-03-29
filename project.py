from comtypes import automation
from comtypes import client
import ctypes
from comtypes import npsupport
import numpy as np

npsupport.enable()

#Set the Operation Object
os = client.GetActiveObject("StaadPro.OpenSTAAD")
geometry = os.geometry
load = os.load
output = os.output
view = os.view
table = os.table

#Define the function for create the array from variant
def make_safe_array_double(size):
    return automation._midlSAFEARRAY(ctypes.c_double).create([0]*size)
def make_safe_array_int(size):
    return automation._midlSAFEARRAY(ctypes.c_int).create([0]*size)
def make_safe_array_long(size):
    return automation._midlSAFEARRAY(ctypes.c_long).create([0]*size)
def make_safe_array_string(size) :
    return automation._midlSAFEARRAY(automation.BSTR).create([""]*size)

def make_variant_vt_ref(obj, var_type):
    var = automation.VARIANT()
    var._.c_void_p = ctypes.addressof(obj)
    var.vt = var_type | automation.VT_BYREF
    return var

#Flag the Method in case using the First time to run the api
os._FlagAsMethod("GetBaseUnit")

geometry._FlagAsMethod("GetMemberCount")
geometry._FlagAsMethod("GetBeamList")

geometry._FlagAsMethod("GetGroupCount")
geometry._FlagAsMethod("GetGroupNames")

load._FlagAsMethod("GetPrimaryLoadCaseCount")
load._FlagAsMethod("GetPrimaryLoadCaseNumbers")

load._FlagAsMethod("GetLoadCombinationCaseCount")
load._FlagAsMethod("GetLoadCombinationCaseNumbers")

output._FlagAsMethod("GetMinMaxBendingMoment")

view._FlagAsMethod("SelectGroup")

geometry._FlagAsMethod("GetNoOfSelectedBeams")
geometry._FlagAsMethod("GetSelectedBeams")
geometry._FlagAsMethod("ClearMemberSelection")

table._FlagAsMethod("CreateReport")
table._FlagAsMethod("AddTable")
table._FlagAsMethod("SetColumnHeader")
table._FlagAsMethod("SetColumnUnitString")
table._FlagAsMethod("SetCellValue")


def main():
    #print("Base Unit:", os.GetBaseUnit())
    allloads = AllLoads()
    #print(allloads)
    allgroups = AllGroups()
    #print(allgroups)

    Memdata = MaxMoFunc(allloads, allgroups)
    #print(Memdata)
    Mat_Mem = np.array(Memdata)
    MakeTable(Mat_Mem)

#The AllLoads Function will take nothing as input but return all loads list number (primary and combination)
def AllLoads():
    #This function will get the loadcase and loadcombination from the model then combine into one variable
    #Primary Load Case
    TLCase = load.GetPrimaryLoadCaseCount()
    safe_array_load_list = make_safe_array_long(TLCase)
    LCaseList = make_variant_vt_ref(safe_array_load_list, automation.VT_ARRAY | automation.VT_I4)
    load.GetPrimaryLoadCaseNumbers(LCaseList)

    #Combination Load Case
    TLComb = load.GetLoadCombinationCaseCount()
    safe_array_comb_list = make_safe_array_long(TLComb)
    LCombList = make_variant_vt_ref(safe_array_comb_list, automation.VT_ARRAY | automation.VT_I4)
    load.GetLoadCombinationCaseNumbers(LCombList)

    #Combine all load together
    allloads = LCaseList[0] + LCombList[0]
    return allloads

#The AllGrups Function will take nothing as input but return all Groups name(String) in the model
def AllGroups():
    #Count the beam groups in the model
    TLGroup = geometry.GetGroupCount(2)
    safe_array_group_list = make_safe_array_string(TLGroup)
    GNameList = make_variant_vt_ref(safe_array_group_list, automation.VT_ARRAY | automation.VT_BSTR)
    Gname = geometry.GetGroupNames(2, GNameList)
    return list(GNameList[0])

#The MaxMo Function will take All Group list(AllGroups) and All Load list(AllLoadList) as input but return the data of the Maximum Moment
def MaxMoFunc(AllLoadList, AllGroups):

    safe_array_dMin= make_safe_array_double(1)
    dMin = make_variant_vt_ref(safe_array_dMin, automation.VT_R8)

    safe_array_dMinPos= make_safe_array_double(1)
    dMinPos = make_variant_vt_ref(safe_array_dMinPos, automation.VT_R8)

    safe_array_dMax= make_safe_array_double(1)
    dMax = make_variant_vt_ref(safe_array_dMax, automation.VT_R8)

    safe_array_dMaxPos= make_safe_array_double(1)
    dMaxPos = make_variant_vt_ref(safe_array_dMaxPos, automation.VT_R8)

    MomentData = []

    for k, group in enumerate(AllGroups):
        view.SelectGroup(group)
        NoSB = geometry.GetNoOfSelectedBeams()
        safe_array_SB_list = make_safe_array_long(NoSB)
        SBList = make_variant_vt_ref(safe_array_SB_list, automation.VT_ARRAY | automation.VT_I4)
        geometry.GetSelectedBeams(SBList, 1)
        geometry.ClearMemberSelection()

        if len(MomentData) <= k:
            MomentData.append([None, None, None, None])

        for i, SB in enumerate(SBList[0]):
            #print(SB)
            for j, loads in enumerate(AllLoadList):
                output.GetMinMaxBendingMoment(SB, "MZ", loads, dMin, dMinPos, dMax, dMaxPos)
                #print(loads, dMin[0], dMinPos[0], dMax[0], dMaxPos[0])
                MaxMo = max(abs(dMax[0]), abs(dMin[0]))
                #print("k :", k)
                if MomentData[k][1] is None or MomentData[k][1] < MaxMo:
                    MomentData[k][0] = SB
                    MomentData[k][2] = loads
                    MomentData[k][1] = round(MaxMo, 3)
                    MomentData[k][3] = group

            #print(MaxMo)
    return MomentData

#This function take a list variable (ListData) as input then print the table
def MakeTable(ListData):
    SheetNo1 = table.CreateReport(" Absolute Maximum Moment in Group ")
    TableNo1 = table.AddTable(SheetNo1, " Maximum Moment", len(ListData), 4)

    table.SetColumnHeader(SheetNo1, TableNo1, 1, "Member No")
    table.SetColumnUnitString(SheetNo1, TableNo1, 1, " ")

    table.SetColumnHeader(SheetNo1, TableNo1, 2, "Moment")
    table.SetColumnUnitString(SheetNo1, TableNo1, 2, "kip-in")

    table.SetColumnHeader(SheetNo1, TableNo1, 3, "Load No")
    table.SetColumnUnitString(SheetNo1, TableNo1, 3, " ")

    table.SetColumnHeader(SheetNo1, TableNo1, 4, "Group Name")
    table.SetColumnUnitString(SheetNo1, TableNo1, 4, " ")

    #print(ListData)

    for i in range(1, len(ListData)+1):
        for j in range(1, 5):
            table.SetCellValue(SheetNo1, TableNo1, i, j, str(ListData[i-1][j-1]))


if __name__ == "__main__":
    main()
