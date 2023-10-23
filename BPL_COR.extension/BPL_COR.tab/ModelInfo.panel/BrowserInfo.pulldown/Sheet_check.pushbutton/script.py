# -*- coding: utf-8 -*-
__title__   = "Sheets"
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

# Initialize counters
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
viewplans = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Views).WhereElementIsNotElementType().ToElements()
schedules = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Schedules).WhereElementIsNotElementType().ToElements()
sheets = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Sheets).WhereElementIsNotElementType().ToElements()


# Collect all ViewPort instances to check if a view is placed on a sheet
viewports = FilteredElementCollector(revit.doc).OfClass(Viewport).ToElements()

# Create a set of all view IDs that are placed on sheets
view_ids_on_sheets = set([vp.ViewId.IntegerValue for vp in viewports])

viewplans = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Views).WhereElementIsNotElementType().ToElements()


for v in sheets:
    view_name = v.Name
    family = v.LookupParameter("Family").AsValueString()
    if v.IsTemplate is False:
        if family != "Legend":
            schedule_type = v.LookupParameter("Family").AsValueString() if v.LookupParameter("Family") else "??"
            schedule_count += 1
            if v.LookupParameter(browser01_param):
                browser01 = v.LookupParameter(browser01_param).AsValueString()
                if browser01 is None:
                    browser01_countbad += 1
                else:
                    browser01_count += 1
            if v.LookupParameter(browser02_param):
                browser02 = v.LookupParameter(browser02_param).AsValueString()
                if browser02 is None:
                    browser02_countbad += 1
                    print("Unsorted Sheets")
                    print("{} - {} - {} - {} - {} - {}".format( v.Id, v.Name, browser01, browser02, family, v.SheetNumber))
                    print("*"*15)
                else:
                    browser02_count += 1


        # print(schedule_type, schedule_name,browser01)
        # Check if schedule is on a sheet
        # if s.Id.IntegerValue in view_ids_on_sheets:
        #     schedule_on_sheet_count += 1
        # else:
        #     schedule_not_on_sheet_count += 1




print("BrowserOrganisation: Sheets")
print(browser01_param, "?? :" ,browser01_countbad)
print(browser01_param,"ok :" ,browser01_count)
print(browser02_param,"?? :" ,browser02_countbad)
print(browser02_param,"ok :" ,browser02_count)
print("Total Sheets; ", schedule_count)