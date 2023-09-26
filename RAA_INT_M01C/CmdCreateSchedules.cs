#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

#endregion

namespace RAA_INT_M01C
{
    [Transaction(TransactionMode.Manual)]
    public class CmdCreateSchedules : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // Get all rooms and get deparment list
            FilteredElementCollector collectorRooms = new FilteredElementCollector(doc);
            collectorRooms.OfCategory(BuiltInCategory.OST_Rooms);
            collectorRooms.WhereElementIsNotElementType();

            // Get parameter name
            List<string> departments = new List<string>() {};

            
            Element roomInst = collectorRooms.FirstElement();
            Parameter roomDepartment = roomInst.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);

            foreach (Element curRoom in collectorRooms)
            {
                Parameter departmentName = curRoom.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);

                string dptName = GetParamValue(curRoom, "Department");
                departments.Add(dptName);
            }

            // Get unique departments
            List<string> uniqueDepartment = departments.Distinct().ToList();
            uniqueDepartment.Sort();


            // Create one schedule per deparment
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Schedules");

                foreach(string name in uniqueDepartment)
                {
                    ElementId catR = new ElementId(BuiltInCategory.OST_Rooms);
                    ViewSchedule curSchedule = ViewSchedule.CreateSchedule(doc, catR);
                    curSchedule.Name = "Dept - " + name;

                    //Parameters                    
                    Element rInst = collectorRooms.FirstElement();

                    Parameter roomNumber = rInst.get_Parameter(BuiltInParameter.ROOM_NUMBER);
                    Parameter roomName = rInst.get_Parameter(BuiltInParameter.ROOM_NAME);
                    Parameter roomDepartmentN = rInst.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                    Parameter roomComments = rInst.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                    Parameter roomLevel = rInst.LookupParameter("Level");
                    Parameter roomArea = rInst.get_Parameter(BuiltInParameter.ROOM_AREA);
                   
                    //Create fields
                    ScheduleField rNumberField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNumber.Id);
                    ScheduleField rNameField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, roomName.Id);
                    ScheduleField rDeptField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, roomDepartmentN.Id);
                    ScheduleField rCommentsField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, roomComments.Id);
                    ScheduleField rLevelField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, roomLevel.Id);
                    ScheduleField rAreaField = curSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomArea.Id);
                    
                    //Hide field (level)
                    rLevelField.IsHidden = true;

                    //Display total (Calculate totals)
                    rAreaField.DisplayType = ScheduleFieldDisplayType.Totals;

                    //Filter by department 
                    ScheduleFilter deptFilter = new ScheduleFilter (rDeptField.FieldId, ScheduleFilterType.Equal, name);
                    curSchedule.Definition.AddFilter(deptFilter);

                    //Group by level and name
                    ScheduleSortGroupField sortByLevel = new ScheduleSortGroupField(rLevelField.FieldId);
                    sortByLevel.ShowHeader = true;
                    sortByLevel.ShowFooter = true;
                    sortByLevel.ShowBlankLine = true;

                    curSchedule.Definition.AddSortGroupField(sortByLevel);

                    ScheduleSortGroupField sortByName = new ScheduleSortGroupField(rNameField.FieldId);
                    curSchedule.Definition.AddSortGroupField(sortByName);
                   

                    //Set totals
                    curSchedule.Definition.IsItemized= true;    
                    curSchedule.Definition.ShowGrandTotal = true;
                    curSchedule.Definition.ShowGrandTotalTitle = true;
                    curSchedule.Definition.ShowGrandTotalCount = true;

                }

                /////////////BONUS////////////////////
                ElementId catRo = new ElementId(BuiltInCategory.OST_Rooms);
                ViewSchedule allDeptSchedule = ViewSchedule.CreateSchedule(doc, catRo);
                allDeptSchedule.Name = "All Departments";

                Element inst = collectorRooms.Last();

                //Parameters
                Parameter roomAllDDepartment = inst.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                Parameter roomAllDArea = inst.get_Parameter(BuiltInParameter.ROOM_AREA);

                //Create fields
                ScheduleField rAllDDeptField = allDeptSchedule.Definition.AddField(ScheduleFieldType.Instance, roomAllDDepartment.Id);
                ScheduleField rAllDAreaField = allDeptSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomAllDArea.Id);

                // Show total area
                rAllDAreaField.DisplayType = ScheduleFieldDisplayType.Totals;

                //Group by department 
                ScheduleSortGroupField sortByDept = new ScheduleSortGroupField(rAllDDeptField.FieldId);
                //sortByDept.ShowHeader = true;
               // sortByDept.ShowFooter = true;
               // sortByDept.ShowBlankLine = true;

                allDeptSchedule.Definition.AddSortGroupField(sortByDept);

                allDeptSchedule.Definition.IsItemized = false;

                //Set totals
                allDeptSchedule.Definition.ShowGrandTotal = true;
                allDeptSchedule.Definition.ShowGrandTotalTitle = true;
                allDeptSchedule.Definition.ShowGrandTotalCount = true;

                tx.Commit();
            }


            // Alert user
            TaskDialog.Show("RAA", $"{uniqueDepartment.Count} individual schedules created, and extra one.");

            return Result.Succeeded;
        }

        private string GetParamValue(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                Debug.Print(curParam.Definition.Name);
                if (curParam.Definition.Name == paramName)
                {
                    Debug.Print(curParam.AsString());
                    return curParam.AsString();
                }
            }

            return null;
        }
    }

}
