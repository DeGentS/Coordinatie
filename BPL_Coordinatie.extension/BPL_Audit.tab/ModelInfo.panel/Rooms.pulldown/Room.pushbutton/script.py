# -*- coding: utf-8 -*-
__title__   = "Rooms info"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je de Rooms/Spaces Info van het project
_____________________________________________________________
Last update:

- [22-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS
import clr
import os
import sys
import System
import shutil


# Importeren van Revit API-elementen
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic
from rpw.ui.forms import SelectFromList

# #----------------------VARIABLES--------------------------------------------------------
# #VARIABLES
#
doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN


levels = FilteredElementCollector(doc).OfCategory(Bic.OST_Levels).WhereElementIsNotElementType().ToElements()
all_rooms = FilteredElementCollector(doc).OfCategory(Bic.OST_Rooms).WhereElementIsNotElementType().ToElements()
all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
all_spaces = FilteredElementCollector(doc).OfCategory(Bic.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()

print("room_id -- room_cat -- room_number -- room_name --  room_level -- room_phase")
for r in all_rooms:
    room_id = r.Id
    room_cat = r.Category.Name
    room_name = r.get_Parameter(BuiltInParameter.ROOM_NAME).AsValueString()
    room_number = r.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsValueString()
    room_level =  r.get_Parameter(BuiltInParameter.LEVEL_NAME).AsValueString()
    room_phase = r.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
    print('{} -- {} - {} -- {} -- {} -- {}'.format(room_id,room_cat,room_number,room_name, room_level,room_phase))



print("space_id --space_cat -- space_number -- space_name --  space_level -- space_phase")
for s in all_spaces:
    space_id = s.Id
    space_cat = s.Category.Name
    space_name = s.get_Parameter(BuiltInParameter.ROOM_NAME).AsValueString()
    space_number = s.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsValueString()
    space_level =  s.get_Parameter(BuiltInParameter.LEVEL_NAME).AsValueString()
    space_phase = s.get_Parameter(BuiltInParameter.ROOM_PHASE).AsValueString()
    print('{} -- {} -- {} -- {} -- {} -- {}'.format(space_id,space_cat,space_number,space_name, space_level, space_phase))


