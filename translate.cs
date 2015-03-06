using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using MyExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ReadGTDocument
{
    interface Translate
    {
        int WriteItnFile(FunctionClass[] functions, string file_name, string sheet_name, ExcelRangeClass excelrange);
    }

    class ExcelTranslater : Translate
    {
        private string _file_name;
        public int GetExcelItemPoints(MyExcel.Worksheet sheet,ExcelRangeClass excelrange)
        {
            if ((excelrange == null) || (sheet == null))  return -1;
            //excelrange.Propname_map
            PropertyInfo[] props = null;
            Type type = typeof(FunctionClass);
            object obj = Activator.CreateInstance(type);
            props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            string strtemp = "";
            bool find;
            foreach (PropertyInfo prop in props)
            {
                find = false;
                for (int i = 0; i < excelrange.F2InCount; i++) 
                {
                    for (int j =0; j < excelrange.Colcount ; j++)
                    {
                        if (sheet.Cells[excelrange.Startrow + i, excelrange.Startcol + j].Value2 != null)
                        {
                            strtemp = sheet.Cells[excelrange.Startrow + i, excelrange.Startcol + j].Value2.ToString(); ;
                            if (strtemp == excelrange.Propname_map[prop.Name])
                            {
                                excelrange.Points.Add(prop.Name, new Point(i, j));                                
                                find = true;
                                break;
                            }
                        }
                    }
                    if (find) break;
                }
            }
            return 0;
        }
        public string File_name
        {
            get { return _file_name; }
            set { _file_name = value; }
        }
       public  int WriteItnFile(FunctionClass[] functions ,string file_name, string sheet_name ,ExcelRangeClass excelrange)
        {
           int GridRowCount = excelrange.Rowcount;
           int F2InCount = excelrange.F2InCount;
           int GridColCount = excelrange.Colcount;
           if (file_name != "")  _file_name  = file_name;
           MyExcel.Application excelApp = new MyExcel.Application();
           try
           { 
               excelApp.Visible = true;
               if (!File.Exists(_file_name)) return -1;

               excelApp.Workbooks.Open(_file_name);
               MyExcel.Worksheet sheet = (MyExcel.Worksheet)excelApp.Worksheets[sheet_name];
               if (excelrange.Points.Count <= 0)
                 GetExcelItemPoints(sheet,excelrange);
               int iStart = excelrange.Startrow;
               int iStartcol = excelrange.Startcol;
               int iInIndex = 0;
               int iOutIndex = 0;
               MyExcel.Range range = (MyExcel.Range)sheet.Range[sheet.Cells[iStart, 1], sheet.Cells[iStart + GridRowCount, iStartcol + GridColCount]];
               int itemp = 0;
               int x, y;
               foreach (FunctionClass function in functions)
               {
                   range.Select();
                   range.Copy();
                   iStart = iStart + GridRowCount + 1;
                   sheet.Rows[iStart].insert();
                   foreach (KeyValuePair<string, Point> kvp in excelrange.Points)
                   {
                      MyExcel.Range rg = (MyExcel.Range)sheet.Range[sheet.Cells[iStart + kvp.Value.X, iStartcol + kvp.Value.Y], sheet.Cells[iStart + kvp.Value.X, iStartcol + kvp.Value.Y]];
                      //rg.Next.Value2 = function.GetType().GetProperty(kvp.Key).GetValue(function, null);
                      bool bl = (bool)(rg.MergeCells);
                      if (bl)
                      {
                          x = rg.MergeArea.Row + rg.MergeArea.Rows.Count;
                          y = rg.MergeArea.Column + rg.MergeArea.Columns.Count;
                          MyExcel.Range rg1 = (MyExcel.Range)sheet.Range[sheet.Cells[x - 1, y], sheet.Cells[x - 1, y]];
                          rg1.Value2 = function.GetType().GetProperty(kvp.Key).GetValue(function, null);
                      }
                      else
                      {
                          rg.Next.Value2 = function.GetType().GetProperty(kvp.Key).GetValue(function, null);
                      }
                   }

                   iInIndex = iStart + F2InCount;
                   MyExcel.Range rangerow = sheet.Range[sheet.Cells[iInIndex, 1], sheet.Cells[iInIndex, iStartcol + GridColCount]];
                   rangerow.Select();
                   rangerow.Copy();                         
                   for (int j = 1; j < function.In_fields.Count(); j++)
                   {
                       sheet.Rows[iInIndex].insert(Shift: (MyExcel.XlDirection.xlDown));
                       rangerow = sheet.Range[sheet.Cells[iInIndex + 1, 1], sheet.Cells[iInIndex + 1, iStartcol + GridColCount]];
                       rangerow.Select();
                       rangerow.Copy();  
                   }
                   itemp = 0;
                   foreach(FieldClass field in function.In_fields)
                   {
                       sheet.Cells[iInIndex + itemp, iStartcol + 1] = field.Name;
                       itemp++;
                   }
                   iOutIndex = iInIndex + function.In_fields.Count() + 1;

                   for (int j = 1; j < function.Out_fields.Count(); j++)
                   {
                       sheet.Rows[iOutIndex].insert(Shift: (MyExcel.XlDirection.xlDown));
                       rangerow = sheet.Range[sheet.Cells[iOutIndex + 1, 1], sheet.Cells[iOutIndex + 1, iStartcol + GridColCount]];
                       rangerow.Select();
                       rangerow.Copy();  
                   }
                   itemp = 0;
                   foreach (FieldClass field in function.Out_fields)
                   {
                       sheet.Cells[iOutIndex + itemp, iStartcol + 1] = field.Name;
                       itemp++;
                   }

                   iStart = iStart + function.In_fields.Count() + function.Out_fields.Count() - 2;
               }
               excelApp.Workbooks[1].RefreshAll();
               MyExcel.Workbook mybook = excelApp.Workbooks[1];
               mybook.Save();
              // mybook.Close(false);
           }
           finally
           {
              // excelApp.Quit();
               GC.Collect();
               excelApp = null;
           }
           return 0;
       }
    }
}
