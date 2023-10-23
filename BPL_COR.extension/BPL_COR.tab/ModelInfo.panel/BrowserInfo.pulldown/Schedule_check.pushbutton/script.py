# -*- coding: utf-8 -*-
__title__   = "Schedule"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 05-12-23
_____________________________________________________________________
Description:

Verkrijg alle views in het project
µ_____________________________________________________________
Last update:

- [05-12-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""
#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic
from pyrevit import revit, DB


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application


browser01_param = "VK_T browser organization"
browser02_param = "VK_T browser organization two"

schedule_on_sheet_count = 0
schedule_not_on_sheet_count = 0

schedule_count = 0
sheet_count = 0
browser01_count = 0
browser01_countbad = 0
browser02_count = 0
browser02_countbad = 0
#----------------------MAIN--------------------------------------------------------
#MAIN


# Get all ViewPlan instances in the document
schedules = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Schedules).WhereElementIsNotElementType().ToElements()
viewplans = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Views).WhereElementIsNotElementType().ToElements()

# Collect all ViewPort instances to check if a view is placed on a sheet
viewports = FilteredElementCollector(revit.doc).OfClass(Viewport).ToElements()

# Create a set of all view IDs that are placed on sheets
view_ids_on_sheets = set([vp.ViewId.IntegerValue for vp in viewports])





for s in schedules:
    schedule_name = s.Name
    if s.IsTemplate is False:

        if "Revision Schedule" not in schedule_name:
            schedule_type = s.LookupParameter("Family").AsValueString() if s.LookupParameter("Family") else "??"
            schedule_count += 1
            if s.LookupParameter(browser01_param):
                browser01 = s.LookupParameter(browser01_param).AsValueString()
                if browser01 is None:
                    browser01_countbad += 1
                else:
                    browser01_count += 1
            if s.LookupParameter(browser02_param):
                browser02 = s.LookupParameter(browser02_param).AsValueString()
                if browser02 is None:
                    browser02_countbad += 1
                else:
                    browser02_count += 1
            # print(s.Id,s.Name,browser01,browser02)

            # print(schedule_type, schedule_name,browser01)
            # Check if schedule is on a sheet
            # if s.Id.IntegerValue in view_ids_on_sheets:
            #     schedule_on_sheet_count += 1
            # else:
            #     schedule_not_on_sheet_count += 1


print("BrowserOrganisation: Schedules")
print(browser01_param, "?? :" ,browser01_countbad)
print(browser01_param,"ok :" ,browser01_count)
print(browser02_param,"?? :" ,browser02_countbad)
print(browser02_param,"ok :" ,browser02_count)
print("Total Schedule; ", schedule_count)
