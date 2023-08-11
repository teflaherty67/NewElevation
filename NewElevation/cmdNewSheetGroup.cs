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
            string excelFile = @"S:\Shared Folders\!RBA Addins\Lifestyle Design\Data Source\NewSheetSetup.xlsx";

            List<List<string>> dataSheets = new List<List<string>>();

            using (var package = new ExcelPackage(excelFile))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelWorkbook wb = package.Workbook;

                ExcelWorksheet ws;

                if (curForm.GetComboboxFoundation() == "Basement" && curForm.GetComboboxFloors() == "One Story")
                    ws = wb.Worksheets[0];
                else if (curForm.GetComboboxFoundation() == "Basement" && curForm.GetComboboxFloors() == "Two Story")
                    ws = wb.Worksheets[1];
                else if (curForm.GetComboboxFoundation() == "Crawlspace" && curForm.GetComboboxFloors() == "One Story")
                    ws = wb.Worksheets[2];
                else if (curForm.GetComboboxFoundation() == "Crawlspace" && curForm.GetComboboxFloors() == "Two Story")
                    ws = wb.Worksheets[3];
                else if (curForm.GetComboboxFoundation() == "Slab" && curForm.GetComboboxFloors() == "One Story")
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
            }

            // create sheets with specifed titleblock           

            foreach (List<string> curSheetData in dataSheets)
            {
                FamilySymbol tBlock = Utils.GetTitleBlockByNameContains(curDoc, curSheetData[2]);

                ElementId tBlockId = tBlock.Id;

                ViewSheet curSheet = ViewSheet.Create(curDoc, tBlockId);

                // add elevation designation to sheet number

                curSheet.SheetNumber = curSheetData[0] + curForm.GetComboboxElevation().ToLower();
                curSheet.Name = curSheetData[1];
            }                

            return Result.Succeeded;
        }
    }
}

