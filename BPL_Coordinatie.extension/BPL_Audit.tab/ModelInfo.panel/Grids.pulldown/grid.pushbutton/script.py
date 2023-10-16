# -*- coding: utf-8 -*-
__title__   = "Grids"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je de info van de aanwezige Grids
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
all_grids = FilteredElementCollector(doc).OfCategory(Bic.OST_Grids).WhereElementIsNotElementType().ToElements()


for g in all_grids:
    grid_id = g.Id
    grid_name = g.get_Parameter(BuiltInParameter.DATUM_TEXT).AsValueString()
    grid_type = g.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM).AsValueString()
    grid_pin = g.Pinned
    print('{} -- {} - {} -- {} '.format(grid_id,grid_name,grid_type,grid_pin))