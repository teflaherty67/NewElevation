#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Controls;
using Forms = System.Windows.Forms;

#endregion

namespace NewElevation
{
    [Transaction(TransactionMode.Manual)]
    public class cmdNewSheetGroup : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document curDoc = uidoc.Document;

            frmNewSheetGroup curForm = new frmNewSheetGroup()
            {
                Width = 320,
                Height = 420,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            curForm.ShowDialog();
            
            // hard-code Excel file
            string excelFile = "S:\\Shared Folders\\!RBA Addins\\Lifestyle Design\\Data Source\\NewSheetSetup.xlsx";

            // create a list to hold the sheetdata
            List<List<string>> dataSheets = new List<List<string>>();

            // get data from the form
            string newElev = curForm.GetComboboxElevation();            

            // set some variables for paramter values

            string newFilter = "";

            if (newElev == "A")
                newFilter = "1";
            else if (newElev == "B")
                newFilter = "2";
            else if (newElev == "C")
                newFilter = "3";
            else if (newElev == "D")
                newFilter = "4";
            else if (newElev == "S")
                newFilter = "5";
            else if (newElev == "T")
                newFilter = "6";

            using (var package = new ExcelPackage(excelFile))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelWorkbook wb = package.Workbook;

                ExcelWorksheet ws;

                if (curForm.GetComboboxFoundation() == "Basement" && curForm.GetComboboxFloors() == "1")
                    ws = wb.Worksheets[0];
                else if (curForm.GetComboboxFoundation() == "Basement" && curForm.GetComboboxFloors() == "2")
                    ws = wb.Worksheets[1];
                else if (curForm.GetComboboxFoundation() == "Crawlspace" && curForm.GetComboboxFloors() == "1")
                    ws = wb.Worksheets[2];
                else if (curForm.GetComboboxFoundation() == "Crawlspace" && curForm.GetComboboxFloors() == "2")
                    ws = wb.Worksheets[3];
                else if (curForm.GetComboboxFoundation() == "Slab" && curForm.GetComboboxFloors() == "1")
                    ws = wb.Worksheets[4];
                else
                    ws = wb.Worksheets[5];

                // get row & column count

                int rows = ws.Dimension.Rows;
                int columns = ws.Dimension.Columns;

                // read Excel data into a list                

                for (int i = 1; i <= rows; i++)
                {
                    List<string> rowData = new List<string>();
                    for (int j = 1; j <= columns; j++)
                    {
                        string cellContent = ws.Cells[i, j].Value.ToString();
                        rowData.Add(cellContent);
                    }
                    dataSheets.Add(rowData);
                }

                dataSheets.RemoveAt(0);
            }

            // create sheets with specifed titleblock
            using(Transaction t = new Transaction(curDoc))
            {
                t.Start("Create Sheets");

                foreach (List<string> curSheetData in dataSheets)
                {
                    FamilySymbol tblock = Utils.GetTitleBlockByNameContains(curDoc, curSheetData[2]);
                    ElementId tBlockId = tblock.Id;

                    ViewSheet curSheet = ViewSheet.Create(curDoc, tBlockId);

                    // add elevation designation to sheet number
                    curSheet.SheetNumber = curSheetData[0] + curForm.GetComboboxElevation().ToLower();
                    curSheet.Name = curSheetData[1];

                    // set parameter values                    
                    Utils.SetParameterByName(curSheet, "Category", "Active");
                    Utils.SetParameterByName(curSheet, "Group", "Elevation " + curForm.GetComboboxElevation());
                    Utils.SetParameterByName(curSheet, "Elevation Designation", curForm.GetComboboxElevation());
                    Utils.SetParameterByName(curSheet, "Code Filter", newFilter);
                    Utils.SetParameterByName(curSheet, "Index Position", int.Parse(curSheetData[3]));
                }

                t.Commit();
            }            

            return Result.Succeeded;
        }
    }
}

