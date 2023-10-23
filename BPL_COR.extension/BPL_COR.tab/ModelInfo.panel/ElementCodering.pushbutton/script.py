# -*- coding: utf-8 -*-
__title__   = "Element\nCodering"
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

    # Iterate through the selected elements
    for ref in selected_ids:
        element = doc.GetElement(ref.ElementId)
        # Use LookupParameter to get specific parameter values
        name_param = element.LookupParameter("Name")
        family_param = element.LookupParameter("Family")
        type_param = element.LookupParameter("Family and Type")
        category = element.LookupParameter('Category')
        nested = element.LookupParameter('SuperComponent')

        # Get parameter values if they exist
        name = name_param.AsString() if name_param and name_param.HasValue else "Unknown"
        family = family_param.AsValueString() if family_param and family_param.HasValue else "Unknown"
        type_name = type_param.AsValueString() if type_param else "Unknown"
        category = category.AsValueString() if category else "Unknown"
        has_nested_value = "Yes" if nested and nested.HasValue else "No"

        # Create a unique key for each element type
        key = (name, family, type_name, category, has_nested_value)

        # Count the occurrences of each element type
        if key in element_summary:
            element_summary[key] += 1
        else:
            element_summary[key] = 1
    print("Name § Family § Type § category § Nested §  Count")

    # Print the summary of elements
    for (name, family, type_name, category, has_nested_value), count in element_summary.items():


        print('{} § {} § {} § {} § {} § {}'.format(name, family, type_name, category, has_nested_value, count))

    # Stop the stopwatch and print elapsed time
    stopwatch.Stop()
    print("\nTime elapsed: {} ms".format(stopwatch.ElapsedMilliseconds))

# Main function
def main():
    # Count and list selected elements in the model
    count_and_list_selected_elements(doc, uidoc)

# Execute the main function
main()