# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import*
from Autodesk.Revit.DB import BuiltInCategory as Bic

doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
#----------------------MAIN--------------------------------------------------------
#MAIN

pbp = FilteredElementCollector(doc).OfCategory(Bic.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
sp = FilteredElementCollector(doc).OfCategory(Bic.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()

def info_pbp(doc):

    info_projectbasepoint = []

    for p in pbp:
        angle = p.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsValueString()
        ew = p.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsValueString()
        ns = p.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsValueString()
        elev = p.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsValueString()

        pbp_info = "NS : {} \n EW : {} \n elev : {} \n angle : {} " .format(ns,ew,elev,angle)
        print("PROJECT_BASE_POINT \n {}".format(pbp_info))
        info_projectbasepoint.append(pbp_info)


    return info_projectbasepoint

def info_svp(doc):

    info_surveypoint = []

    for s in sp:
        sv_ew = s.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsValueString()
        sv_ns = s.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsValueString()
        sv_elev = s.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsValueString()

        sv_info = "NS : {} \n EW : {} \n elev : {} " .format(sv_ns,sv_ew,sv_elev)
        print("Survey_POINT \n {}".format(sv_info))
        info_surveypoint.append(sv_info)


    return info_surveypoint

