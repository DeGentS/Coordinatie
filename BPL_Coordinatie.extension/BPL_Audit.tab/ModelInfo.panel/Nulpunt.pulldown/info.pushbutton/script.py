# -*- coding: utf-8 -*-
__title__   = "Coordinates"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je info van de coordinaten in het model
Project Base Point & Survey Point
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
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import View3D, ViewOrientation3D, XYZ, BoundingBoxXYZ



#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

pbp = FilteredElementCollector(doc).OfCategory(Bic.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
sp = FilteredElementCollector(doc).OfCategory(Bic.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()


for p in pbp:
    angle = p.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsValueString()
    ew = p.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsValueString()
    ns = p.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsValueString()
    elev = p.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsValueString()
    pinned = p.Pinned

    pbp_location = "NS: {} \n EW: {} \n Elev: {} \n Angle: {} \n Pinned: {}".format(ns,ew,elev,angle,pinned)


for s in sp:
    # angles = s.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsValueString()
    ew = s.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsValueString()
    ns = s.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsValueString()
    elev = s.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsValueString()
    clipped = s.Clipped
    pinned = s.Pinned

    sp_location = "NS: {} \n EW: {} \n Elev: {} \n Clipped: {} \n Pinned: {}".format(ns,ew,elev,clipped,pinned)

print(doc.Title)
print(doc.Application.Username)
print(doc.Application.VersionName)
print("{}{}{}".format(5*"-", "PROJECT BASE Point", 5*"-"))
print(pbp_location)
print("{}{}{}".format(5*"-", "SURVEY Point", 5*"-"))
print(sp_location)

