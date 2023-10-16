# -*- coding: utf-8 -*-
__title__   = "Workset export V2"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Hiermee krijg je een overzicht van aanwezige worksets in het model
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


all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()


unique_element_workset = []
category_elements = []

# Loop over alle worksets
for w in all_worksets:
    workset_name = w.Name
    print("Workset:", workset_name)

    # Houd de categorieën voor deze workset bij
    categories_in_workset = []

    # Loop over alle elementen in deze workset
    for e in all_elements:
        workset_element = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()

        # Als het workset-element overeenkomt met de huidige workset
        if workset_element == workset_name:
            category_elem = e.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString()
            categories_in_workset.append(category_elem)

    # Voeg de unieke categorieën voor deze workset toe aan de lijst
    unique_categories = list(set(categories_in_workset))
    category_elements.append((workset_name, unique_categories))

# Print de unieke categorieën voor elk workset
for workset_name, unique_categories in category_elements:
    print("Workset:", workset_name)
    print("Unique Categories:", unique_categories)
