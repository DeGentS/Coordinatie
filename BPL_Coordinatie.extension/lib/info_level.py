# -*- coding: utf-8 -*-


#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

levels = FilteredElementCollector(doc).OfCategory(Bic.OST_Levels).WhereElementIsNotElementType().ToElements()

def level_info(doc):

    list_levels = []

    for level in levels:
        level_id = level.Id
        family_type = level.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
        name = level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsValueString()
        bs = level.get_Parameter(BuiltInParameter.LEVEL_IS_BUILDING_STORY).AsValueString()
        elevation = level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()
        structural = level.get_Parameter(BuiltInParameter.LEVEL_IS_STRUCTURAL).AsValueString()
        ifcname = level.LookupParameter("ifcName")

        list_levels.append((level_id,family_type, name, bs, elevation, ifcname))

        # o = "{} -- {} -- {} -- {} -- {}".format(family_type,name,bs,elevation,ifcname)
        # # print(o)

    print("Level Id -- family_type -- name -- BuildingStory -- elevation -- ifcname")
    # Print sorted output
    for level in list_levels:
        level_id, family_type, name, bs, elevation, ifcname = level
        o = "{} -- {} -- {} -- {} -- {} -- {}".format(level_id, family_type, name, bs, elevation, ifcname)
        print(o)

    return list_levels
