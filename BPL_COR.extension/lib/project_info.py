import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import*
from Autodesk.Revit.DB import BuiltInCategory as Bic

#VARIABLES

doc         = __revit__.ActiveUIDocument.Document
uidoc       = __revit__.ActiveUIDocument
app         = __revit__.Application

#----------------------MAIN--------------------------------------------------------
#MAIN


model = FilteredElementCollector(doc).OfCategory(Bic.OST_ProjectInformation).WhereElementIsNotElementType().ToElements()

def project_info(doc):

    info = []

    for m in model:
        author = m.get_Parameter(BuiltInParameter.PROJECT_AUTHOR).AsString()

        client = m.get_Parameter(BuiltInParameter.CLIENT_NAME).AsString()
        building_name = m.get_Parameter(BuiltInParameter.PROJECT_BUILDING_NAME).AsString()
        project_name = m.get_Parameter(BuiltInParameter.PROJECT_NAME).AsString()
        project_adress = m.get_Parameter(BuiltInParameter.PROJECT_ADDRESS).AsString()
        location = doc.SiteLocation.GeoCoordinateSystemId

        if author is "":
            author = "geen info"

        if client is "":
            client = "Info niet aanwezig"

        if building_name is "":
            building_name = "info niet aanwezig"



        p_info = "Author: {} \n Client: {} \n Building Name: {} \n Project Name: {} \n {} \n {}".format(author,client,building_name,project_name,project_adress,location)
        print(p_info)
        info.append(p_info)



    return info

        # print("Author: {} ".format(author))
        # print("Client: {} ".format(client))
        # print("Building Name: {} ".format(building_name))
        # print("Project Name: {} ".format(project_name))
        # print("Project Adress: {} ".format(project_adress))
        # print("location",location)
