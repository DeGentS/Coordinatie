def browser_schedules(doc, *browser_params):
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory as Bic

    Total_view = 0
    alles = 0
    sorted_schedule = [0] * len(browser_params)
    unsorted_schedule = [0] * len(browser_params)

    schedules = FilteredElementCollector(doc).OfCategory(Bic.OST_Schedules).WhereElementIsNotElementType().ToElements()

    schedule_data = []  # List to store data for views with missing parameter values

    for s in schedules:
        schedule_name = s.Name
        schedule_id = s.Id

        if "Revision Schedule" not in schedule_name:
            schedule_entry = {"schedule Name": schedule_name,"schedule Id": schedule_id}
            alles += 1
            missing_value = False  # Flag to track if any parameter is missing

            for i, browser_param in enumerate(browser_params):
                if s.LookupParameter(browser_param):
                    browser_value = s.LookupParameter(browser_param).AsValueString()
                    if browser_value is None:
                        unsorted_schedule[i] += 1
                        missing_value = True
                        schedule_entry[browser_param] = "None"

                    else:
                        sorted_schedule[i] += 1
                        schedule_entry[browser_param] = browser_value

            if missing_value:
                schedule_data.append(schedule_entry)
                Total_view += 1
    print("Totaal aantal schedules: {}".format(alles))
    print("sorted_schedule: {}".format(sorted_schedule))
    print("unsorted_schedule: {}".format(unsorted_schedule))
    return {
        "nietgesorteerd_view": Total_view,  # Count of views with at least one missing parameter
        "sorted_schedule": sorted_schedule,
        "unsorted_schedule": unsorted_schedule,
        "schedule_data": schedule_data # Collected view data with missing parameter values
    }
