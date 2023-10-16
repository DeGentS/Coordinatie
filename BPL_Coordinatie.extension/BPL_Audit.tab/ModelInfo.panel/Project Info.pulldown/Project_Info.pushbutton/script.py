# -*- coding: utf-8 -*-
__title__   = "Project info"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je de Project Info van het project
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


clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import BuiltInCategory as Bic

#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN


model = FilteredElementCollector(doc).OfCategory(Bic.OST_ProjectInformation).WhereElementIsNotElementType().ToElements()

def get_model_info(m):

    author = m.get_Parameter(BuiltInParameter.PROJECT_AUTHOR).AsString()
    client = m.get_Parameter(BuiltInParameter.CLIENT_NAME).AsString()
    building_name = m.get_Parameter(BuiltInParameter.PROJECT_BUILDING_NAME).AsString()
    project_name = m.get_Parameter(BuiltInParameter.PROJECT_NAME).AsString()
    project_number = m.get_Parameter(BuiltInParameter.PROJECT_NUMBER).AsString()
    project_adress = m.get_Parameter(BuiltInParameter.PROJECT_ADDRESS).AsString()
    location = doc.SiteLocation.GeoCoordinateSystemId

    return(author, client,building_name,project_name,project_number,project_adress,location)

for m in model:
    model_info = get_model_info(m)
    author, client, building_name, project_name, project_number, project_adress, location = model_info
    print(" Author: {} \n Client: {} \n Building Name: {} \n Project Name: {} \n Project Number: {} \n Project Adress : {} \n Location: {}".format(author, client, building_name, project_name, project_number, project_adress, location))