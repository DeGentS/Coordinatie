#-----------------------IMPORTS-------------------------------------------------------
#IMPORTS


import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import*
from Autodesk.Revit.DB import BuiltInCategory as Bic
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter, ElementId, Options


#----------------------VARIABLES--------------------------------------------------------
#VARIABLES

doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN

def phasing_overview(doc):

    phasing = FilteredElementCollector(doc).OfCategory(Bic.OST_Phases).WhereElementIsNotElementType().ToElements()


    for phase in phasing:
        print(phase.Name)
        # print(phase.PhaseFilter)

    # Initialisatie van een dictionary om het aantal elementen per fase bij te houden
    elements_per_phase = {}

    # Maak een FilteredElementCollector voor alle elementen met geometrie
    collector = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

    # Loop door elk element
    for element in collector:
        # Haal de 'Phase Created'-parameter op
        phase_created_param = element.get_Parameter(BuiltInParameter.PHASE_CREATED)

        # Controleer of het element een geldige 'Phase Created'-parameter heeft
        if phase_created_param and phase_created_param.AsElementId() != ElementId.InvalidElementId:
            phase_id = phase_created_param.AsElementId()

            # Update de telling voor de fase
            if phase_id in elements_per_phase:
                elements_per_phase[phase_id] += 1
            else:
                elements_per_phase[phase_id] = 1

    # Print de resultaten
    for phase_id, count in elements_per_phase.items():
        phase = doc.GetElement(phase_id)
        if phase:  # Controleer of de fase bestaat
            print("Fase Naam: {}, Aantal Elementen: {}".format(phase.Name, count))