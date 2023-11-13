using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;


namespace DC.DataLogger
{
    public class ExcelDataLogger<T> : DataLoggerBase, IDataLogger<T>
    {

        private int _lastRow;
        private static ExcelManger _excel;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name">file name</param>
        /// <param name="path">optional parameter
        /// if empty file will be written to base directory
        /// </param>
        public ExcelDataLogger(string name, string path = "") : base(name, "xlsx",path)
        {

           // _excel = new ExcelManger(LoggersPath);

        }



        public static ExcelManger ExcelManger
        {
            get
            {
                if (_excel == null)
                    _excel = new ExcelManger(LoggersPath,true);
                return _excel;
            }
        }

        protected override void Initialize()
        {
            

            if (!File.Exists(LoggersPath))               
                CreateNewFile();

        }

        private void CreateNewFile()
        {
            ExcelManger.AddWorkBook();
           /* _workBook = _excelApp.Workbooks.Add();
            _xlWorksheet = _workBook.Worksheets[_workBook.Worksheets.Count];
            _xlRange = _xlWorksheet.UsedRange;*/
          //  _lastRow = 1;
            SaveAs();
           // WriteHeaders();
           // _workBook.Close(0);
        }

        public List<T> GetEntryIds(string path)// where T:new()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
           
            Workbook xlWorkbook = _excel.Application.Workbooks.Open(path);
            _Worksheet xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            List<T> rows = new List<T>();

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 2; i <= rowCount; i++)
            {
                DateTimeOffset dateTimeOffset = DateTimeOffset.FromUnixTimeSeconds(long.Parse(xlRange.Cells[i, 2].ToString()));
                // DateTimeOffset dateTimeOffset2 = DateTimeOffset.FromUnixTimeMilliseconds(epochMilliseconds);

                /* if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                     rows.Add(new T()
                     {
                         EntryId = xlRange.Cells[i, 1].Value2.ToString(),
                        CreationDate = dateTimeOffset.DateTime,
                         Views = int.Parse(xlRange.Cells[i, 4].Value2.ToString())
                     });*/

            }

            _excel.Close();

            return rows;
        }

        public List<T> GetEntryIds<T>(string v)
        {
            throw new NotImplementedException();
        }
        private IList<PropertyInfo> ColumnNames
        {
            get
            {
                Type myType = typeof(T);
                IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

                return props;
            }
        }

      

        private void SaveAs()
        {
            if (!File.Exists(LoggersPath))
                ExcelManger.WorkBook.SaveAs(LoggersPath, XlFileFormat.xlOpenXMLWorkbook);
            
        }

        public void Add(T item)
        {
            ExcelManger.Open(LoggersPath);
            int j = 1;

            foreach (PropertyInfo prop in ColumnNames)
            {
                var val = prop.GetValue(item, null);
                try
                {
                    if (val != null)
                        ExcelManger.XlWorksheet.Cells[_lastRow + 1, j] = val.ToString();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                j++;
            }

            _lastRow++;


            ExcelManger.WorkBook.Save();
            ExcelManger.WorkBook.Close(0);          
        }


        /* public void SaveListToExcel(List<T> data)
         {
             Console.Clear();

             Console.WriteLine("Writing excel");
             //string path = Console.ReadLine();
             string path = _filename;// ;
             string filePath = @"" + path;

             //Application excel = new Application();
             //Workbook wb = excel.Workbooks.Add();
             Worksheet ws = _workBook.Worksheets[1];
             _workBook.SaveAs(path, XlFileFormat.xlOpenXMLWorkbook);
             try
             {

                 Type myType = typeof(T);
                 IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());
                 int i = 2, j = 1;

                 //Write headers
                 foreach (PropertyInfo prop in props)
                 {
                     ws.Cells[1, j].Value2 = prop.Name;
                     j++;
                 }
                 j = 1;

                 foreach (var row in data)
                 {

                     foreach (PropertyInfo prop in props)
                     {
                         ws.Cells[i, j].Value2 = prop.GetValue(row, null);
                         j++;

                     }
                     j = 1;
                     i++;
                 }
                 _workBook.Save();

             }
             catch (Exception ex)
             {
                 Console.WriteLine(ex.Message);
             }
             finally
             {
                 _workBook.Close(0);
                 FinalizeExcel(_workBook, ws, _excelApp);
             }
         }
        */
       
       

        public void SaveList(List<T> data)
        {
            throw new NotImplementedException();
        }

      

        protected override void WriteHeaders()
        {
            int j = 1;
            //Write headers
            foreach (var prop in ColumnNames)
            {
                ExcelManger.XlWorksheet.Cells[_lastRow, j] = prop.Name;
                j++;
            }
            _lastRow++;
            ExcelManger.WorkBook.Save();
          
        }

        public void Close()
        {
            ExcelManger.Close();
        }
    }

  
    

}
