# -*- coding: utf-8 -*-
__title__   = "Info"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Krijg snel een overzicht van de model info.
_____________________________________________________________
Last update:

- [22-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS
import clr

# Importeren van Revit API-elementen
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic

# Custom IMPORTS

from PBP import info_pbp as ip
from PBP import info_svp as sp
from project_info import project_info as pi
from info_level import level_info as li

#
# #----------------------VARIABLES--------------------------------------------------------
# #VARIABLES
#
doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

print(doc.Title)
print("checked by: {}".format(doc.Application.Username))
print(doc.Application.VersionName)
print(10*"-")
project = pi(doc)
print(10*"-")
levels = li(doc)
print(10*"-")
info = ip(doc)

survey = sp(doc)
