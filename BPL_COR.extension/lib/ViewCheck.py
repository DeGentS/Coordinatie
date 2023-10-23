def view_plans(doc, *browser_params):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory as Bic

    Total_view = 0
    alles = 0
    sorted_view = [0] * len(browser_params)
    unsorted_view = [0] * len(browser_params)

    views = FilteredElementCollector(doc).OfCategory(Bic.OST_Views).WhereElementIsNotElementType().ToElements()

    view_data = []  # List to store data for views with missing parameter values

    for v in views:
        view_name = v.Name
        view_id = v.Id
        family = v.LookupParameter("Family").AsValueString() if v.LookupParameter("Family") else "??"

        if not v.IsTemplate and family != "Legend":
            view_entry = {"View Name": view_name,"View Id": view_id}
            alles += 1
            missing_value = False  # Flag to track if any parameter is missing

            for i, browser_param in enumerate(browser_params):
                if v.LookupParameter(browser_param):
                    browser_value = v.LookupParameter(browser_param).AsValueString()
                    if browser_value is None:
                        unsorted_view[i] += 1
                        missing_value = True
                        view_entry[browser_param] = "None"

                    else:
                        sorted_view[i] += 1
                        view_entry[browser_param] = browser_value

            if missing_value:
                view_data.append(view_entry)
                Total_view += 1

    print("Totaal aantal views: {}".format(alles))
    print("sorted_view: {}".format(sorted_view))
    print("unsorted_view: {}".format(unsorted_view))
    return {
        "nietgesorteerd_view": Total_view,  # Count of views with at least one missing parameter
        "sorted_view": sorted_view,
        "unsorted_view": unsorted_view,
        "view_data": view_data # Collected view data with missing parameter values
    }
