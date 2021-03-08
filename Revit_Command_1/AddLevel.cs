using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

using Excel = Microsoft.Office.Interop.Excel;



namespace Revit_Command_1
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        //Lists для строк из Excel файла.
        List<string> nameLevels = new List<string>();
        List<string> elevationZ = new List<string>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            //Получение расположения Excel файла.
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "(*.xlsx)|*.xlsx";
            openFile.ShowDialog();

            string fullPath = openFile.FileName;

            ExportExcel(fullPath);

            //Транзация
            using (Transaction tx = new Transaction(doc))
            {
                try
                {
                    for (int i = 0; i < (nameLevels.Count - 1); i++)
                    {
                        //Создать Level из Lists.
                        tx.Start("Create: " + nameLevels[i + 1]);
                        double elevation = Convert.ToInt32(elevationZ[i+1]);
                        Level level = Level.Create(doc, elevation);
                        level.Name = nameLevels[i+1];
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Revit", (ex).ToString());
                    tx.RollBack();
                }
            }
            return Result.Succeeded;
        }
        
        public void ExportExcel(string excelPath)
        {
            try 
            {
                //Открыть и загрузить Excel файл.
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(excelPath);
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                int lastRow = (int)lastCell.Row;

                //Загрузить колонки в List
                for (int i = 0; i < lastRow; i++)
                {
                    nameLevels.Add(ObjWorkSheet.Cells[i + 1, 1].Text.ToString());
                }
                elevationZ.Add(ObjWorkSheet.Cells[1, 2].Text.ToString());
                for (int i = 1; i < lastRow; i++)
                {
                    elevationZ.Add(ObjWorkSheet.Cells[i + 1, 2].Text.ToString());
                }

                ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
                ObjWorkExcel.Quit(); // выйти из Excel
                GC.Collect(); // убрать за собой
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Revit", (ex).ToString());
            }
        }
    }
}