# -*- coding: utf-8 -*-
__title__   = "Views"
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
browser03_param = "VK_T browser organization two"


# schedule_on_sheet_count = 0
# schedule_not_on_sheet_count = 0
#
# schedule_count = 0
# sheet_count = 0
# browser01_count = 0
# browser01_countbad = 0
# browser02_count = 0
# browser02_countbad = 0
# #----------------------MAIN--------------------------------------------------------
# #MAIN
#
#
# # Get all ViewPlan instances in the document
# viewplans = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Views).WhereElementIsNotElementType().ToElements()
# schedules = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Schedules).WhereElementIsNotElementType().ToElements()
# sheets = FilteredElementCollector(revit.doc).OfCategory(Bic.OST_Sheets).WhereElementIsNotElementType().ToElements()
#
#
# # Collect all ViewPort instances to check if a view is placed on a sheet
# viewports = FilteredElementCollector(revit.doc).OfClass(Viewport).ToElements()
#
# # Create a set of all view IDs that are placed on sheets
# view_ids_on_sheets = set([vp.ViewId.IntegerValue for vp in viewports])
#
#
# for v in viewplans:
#     view_name = v.Name
#     family = v.LookupParameter("Family").AsValueString()
#     if v.IsTemplate is False:
#         if family != "Legend":
#             schedule_type = v.LookupParameter("Family").AsValueString() if v.LookupParameter("Family") else "??"
#             schedule_count += 1
#             if v.LookupParameter(browser01_param):
#                 browser01 = v.LookupParameter(browser01_param).AsValueString()
#                 if browser01 is None:
#                     browser01_countbad += 1
#                 else:
#                     browser01_count += 1
#             if v.LookupParameter(browser02_param):
#                 browser02 = v.LookupParameter(browser02_param).AsValueString()
#                 if browser02 is None:
#                     browser02_countbad += 1
#                     print(v.Id, v.Name, browser01, browser02, family)
#                 else:
#                     browser02_count += 1
#
#
#         # print(schedule_type, schedule_name,browser01)
#         # Check if schedule is on a sheet
#         # if s.Id.IntegerValue in view_ids_on_sheets:
#         #     schedule_on_sheet_count += 1
#         # else:
#         #     schedule_not_on_sheet_count += 1
#
#
#
#
# print("BrowserOrganisation: Views")
# print(browser01_param, "?? :" ,browser01_countbad)
# print(browser01_param,"ok :" ,browser01_count)
# print(browser02_param,"?? :" ,browser02_countbad)
# print(browser02_param,"ok :" ,browser02_count)
# print("Total Schedule; ", schedule_count)

def view_plans(doc, *browser_params):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory as Bic

    schedule_count = 0
    browser_counts = [0] * len(browser_params)
    browser_counts_bad = [0] * len(browser_params)

    views = FilteredElementCollector(doc).OfCategory(Bic.OST_Views).WhereElementIsNotElementType().ToElements()

    for v in views:
        view_name = v.Name
        family = v.LookupParameter("Family").AsValueString() if v.LookupParameter("Family") else "??"

        if not v.IsTemplate and family != "Legend":
            schedule_count += 1

            for i, browser_param in enumerate(browser_params):
                if v.LookupParameter(browser_param):
                    browser_value = v.LookupParameter(browser_param).AsValueString()
                    if browser_value is None:
                        browser_counts_bad[i] += 1
                        print(view_name)
                    else:
                        browser_counts[i] += 1

    return {
        "schedule_count": schedule_count,
        "browser_counts": browser_counts,
        "browser_counts_bad": browser_counts_bad
    }

# Example usage
# result = view_plans(doc, browser01_param)
# or
result = view_plans(doc, browser01_param, browser02_param)
# or
# result = view_plans(doc, browser01_param, browser02_param, browser03_param)

print(result)
