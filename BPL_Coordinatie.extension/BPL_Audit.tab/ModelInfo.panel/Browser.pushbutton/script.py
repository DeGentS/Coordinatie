# -*- coding: utf-8 -*-
__title__   = "Browser"
__author__     = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je de info van de browser organisatie
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
from pyrevit import revit, DB


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

# doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

# # Get the active document
# doc = revit.doc
#
# # Get the browser organization for views
# browser_org = BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc)
#
# # Print the browser organization
# print("Project Browser View Organization:")
# print("Name: " + browser_org.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString())
# print("Type: " + str(browser_org.Type))
#
# # Get the sort and group configuration
# sort_group_config = browser_org.GetSortGroupConfiguration()
# print(sort_group_config)

# Get the active document
doc = revit.doc

# # Get the current browser organization for views
# browser_org = BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc)
#
#
# # Inspect the available methods and attributes
# print(dir(browser_org))
#
# # Get the active document
# doc = revit.doc

# Get the current browser organization for views
browser_org = BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc)
browser_org02 = BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc)
browser_schedules = BrowserOrganization.GetCurrentBrowserOrganizationForSchedules(doc)
browser_sheets = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc)

# Get the sorting parameter ID
sorting_param_id = browser_org.SortingParameterId

# Get the parameter from the document
sorting_param = doc.GetElement(sorting_param_id)



# Print the name of the sorting parameter
if sorting_param:
    print("Sorting Parameter Name: " + sorting_param.Name)
else:
    print("No sorting parameter set.")


print(browser_org02.SortingOrder)
print("Views Name: " + browser_org02.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString())
print("Sheets Name: " + browser_sheets.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString())
print("Schedules Name: " + browser_schedules.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString())