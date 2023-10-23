#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS
import clr


# Importeren van Revit API-elementen
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,RevitLinkInstance


# #----------------------VARIABLES--------------------------------------------------------
# #VARIABLES
#
doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN


def linked_model(doc):
    # Collect all Revit Link Instances in the model
    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)

    # Iterate through each link instance and print details
    for link in link_instances:
        link_name = link.Name
        pinned = link.Pinned
        workset = link.LookupParameter("Workset").AsValueString()
        print("**Link Name:** {} - Pinnend: {} - workset: {}".format(link_name, pinned, workset))