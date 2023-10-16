# -*- coding: utf-8 -*-
__title__   = "Phasing"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je info van Phasing in het model
_____________________________________________________________
Last update:

- [22-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS


import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import*
from Autodesk.Revit.DB import BuiltInCategory as Bic

#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

from pyrevit import revit, DB


phasing = FilteredElementCollector(doc).OfCategory(Bic.OST_Phases).WhereElementIsNotElementType().ToElements()

for phase in phasing:
    print(phase.Name)
    # print(phase.PhaseFilter)
