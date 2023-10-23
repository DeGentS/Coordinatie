# -*- coding: utf-8 -*-
__title__   = "Element\nParameter"
__doc__     = """ """
__doc__ = """Version = 1.0
Date    = 05-12-23
_____________________________________________________________________
Description:

Selecteer elementen en verkrijg de info over de benaming.

_____________________________________________________________
Last update:

- [05-12-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan
"""

#§§§§§§§§§§§-IMPORTS§§§§§§§§§§§§§§§§§§§§§§§§§§§-


import clr


# clr.AddReference('RevitAPIUI')
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector, CategoryType, BuiltInCategory, ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from System.Diagnostics import Stopwatch


#§§§§§§§§§§§VARIABLES§§§§§§§§§§§§§§§§§§§§§§§§§§§§
#VARIABLES

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

#§§§§§§§§§§§MAIN§§§§§§§§§§§§§§§§§§§§§§§§§§§§

#§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

# Function to count and list selected elements used in the model
def count_and_list_selected_elements(doc, uidoc):
    # Prompt user to select elements
    selected_ids = uidoc.Selection.PickObjects(ObjectType.Element, "Please select elements")
    if not selected_ids:
        print("No elements selected.")
        return

    # Start a stopwatch for performance measurement
    stopwatch = Stopwatch()
    stopwatch.Start()

    # Create a dictionary to count instances of each element type
    element_summary = {}
    print("ID § Unique revit id § Family and Type § category § Nested §  Keynote")
    # Iterate through the selected elements
    for ref in selected_ids:
        element = doc.GetElement(ref.ElementId)
        element_type = doc.GetElement(element.GetTypeId())
        if element_type is not None:  # Controleer of element_type niet None is
            # Use LookupParameter to get specific parameter values

            id = element.Id
            # revit_uniquer_id = element.LookupParameter('Revit unique id')
            family_param = element.LookupParameter("Family and Type")
            category = element.LookupParameter('Category').AsValueString()
            nested = "Parameter bestaat niet"
            if getattr(element, 'SuperComponent', None):
                nested = "Yes"
            else:
                nested = "No"
            keynote = element_type.LookupParameter('Keynote')
            phase = element.LookupParameter('Phase Created')


        # Get parameter values if they exist
        family = family_param.AsValueString() if family_param and family_param.HasValue else "Unknown"
        # revit_uniquer_id_value = revit_uniquer_id.AsValueString()
        keynote_value = keynote.AsValueString()
        phase_value = phase.AsValueString() if phase and phase.HasValue else "unknown"
        # category = category.AsValueString() or category.AsString() if category else "Unknown"




        print('{} § {} § {} § {} § {} § {}'.format(id, family, category, nested, keynote_value, phase_value))

    # Stop the stopwatch and print elapsed time
    stopwatch.Stop()
    print("\nTime elapsed: {} ms".format(stopwatch.ElapsedMilliseconds))

# Main function
def main():
    # Count and list selected elements in the model
    count_and_list_selected_elements(doc, uidoc)


# Execute the main function
main()