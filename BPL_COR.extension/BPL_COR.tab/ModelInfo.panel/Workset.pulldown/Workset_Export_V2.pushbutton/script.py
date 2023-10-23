# -*- coding: utf-8 -*-
__title__   = "Workset Test"
__author__ = "Sean De Gent"
__doc__ = """Version = 1.0
Date    = 22-09-23
_____________________________________________________________________
Description:

Test routine voor het optimaliseren van de export

_____________________________________________________________
Last update:

- [22-09-23] 1.0 RELEASE


author  = Sean De Gent i.o.v. BimPlan

_____________________________________________________________________
"""

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector, FilteredWorksetCollector, WorksetKind, BuiltInParameter

doc = __revit__.ActiveUIDocument.Document

# Verzamel alle worksets en elementen
all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
all_elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

# Een dictionary om elementen te groeperen op basis van worksets
workset_categories = {}

# Loop door alle elementen
for e in all_elements:
    workset_id = e.WorksetId.IntegerValue
    matching_worksets = [w for w in all_worksets if w.Id.IntegerValue == workset_id]

    if matching_worksets:
        workset = matching_worksets[0]
        workset_name = workset.Name
        category_name = e.Category.Name if e.Category else "No Category"

        if workset_name not in workset_categories:
            workset_categories[workset_name] = set()

        workset_categories[workset_name].add(category_name)

    else:
        print("No matching workset found for WorksetId {}. Skipping.".format(workset_id))

# Print de resultaten
for workset, categories in workset_categories.items():
    print("Workset: {}".format(workset))
    print("Unique Categories: {}".format(list(categories)))