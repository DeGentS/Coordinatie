def browser_sheets(doc, *browser_params):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory as Bic

    Total_view = 0
    alles = 0
    sorted_sheet = [0] * len(browser_params)
    unsorted_sheet = [0] * len(browser_params)

    sheets = FilteredElementCollector(doc).OfCategory(Bic.OST_Sheets).WhereElementIsNotElementType().ToElements()

    sheet_data = []  # List to store data for views with missing parameter values

    for s in sheets:
        sheet_name = s.Name
        sheet_id = s.Id
        sheet_number = s.SheetNumber
        family = s.LookupParameter("Family").AsValueString() if s.LookupParameter("Family") else "??"

        if not s.IsTemplate and family != "Legend":
            sheet_entry = {"sheet Name": sheet_name,"sheet Id": sheet_id, "sheet Number": sheet_number}
            alles += 1
            missing_value = False  # Flag to track if any parameter is missing

            for i, browser_param in enumerate(browser_params):
                if s.LookupParameter(browser_param):
                    browser_value = s.LookupParameter(browser_param).AsValueString()
                    if browser_value is None:
                        unsorted_sheet[i] += 1
                        missing_value = True
                        sheet_entry[browser_param] = "None"

                    else:
                        sorted_sheet[i] += 1
                        sheet_entry[browser_param] = browser_value

            if missing_value:
                sheet_data.append(sheet_entry)
                Total_view += 1
    print("Totaa aantal sheets: {}".format(alles))
    print("sorted_sheet: {} ".format(sorted_sheet))
    print("unsorted_sheet: {}".format(unsorted_sheet))
    return {
        "nietgesorteerd_view": Total_view,  # Count of views with at least one missing parameter
        "sorted_sheet": sorted_sheet,
        "unsorted_sheet": unsorted_sheet,
        "sheet_data": sheet_data # Collected view data with missing parameter values
    }
