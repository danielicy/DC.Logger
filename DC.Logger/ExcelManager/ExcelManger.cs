using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
 
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace DC.DataLogger
{
    public class ExcelManger
    {

        private Application _excelApp;
        private Workbook _workBook;
        private Worksheet _xlWorksheet;
        private Range _xlRange;
        private string _path;

        public ExcelManger(string path, bool init = false)
        {
            if (init)
                Initialize();

            _path = path;
        }

        public Application Application { get => _excelApp; private set => _excelApp = value; }
        public Workbook WorkBook { get => _workBook; private set => _workBook = value; }
        public Worksheet XlWorksheet { get => _xlWorksheet; private set => _xlWorksheet = value; }
        public Range XlRange { get => _xlRange; private set => _xlRange = value; }



        private void Initialize()
        {
            _excelApp = new Application();

        }

        public void AddWorkBook()
        {
            _workBook = Application.Workbooks.Add();
            _xlWorksheet = (Worksheet)WorkBook.Worksheets[WorkBook.Worksheets.Count];
            _xlRange = XlWorksheet.UsedRange;

           // _workBook.Close(0);
        }

        public void AddItem<T>(T item)
        {

        }

        public void AddRange<T>(List<T> items)
        {

        }

        public void CreateWorkBook(string path)
        {

        }

        public static System.Data.DataTable GetWorkSheet(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;


                reader = ExcelReaderFactory.CreateReader(stream);
                //// reader.IsFirstRowAsColumnNames
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);


                return dataSet.Tables[0];


            }
        }
        public void OpenWorkBook()
        {

            _workBook = Application.Application.Workbooks.Open(_path);
            _xlWorksheet = (Worksheet)_workBook.Sheets[1];
            Range xlRange = _xlWorksheet.UsedRange;

        }

        public void OpenWorksheet(int id)
        {
            _workBook = Application.Application.Workbooks.Open(_path);
            _xlWorksheet = (Worksheet)_workBook.Sheets[id];
            _xlRange = _xlWorksheet.UsedRange;

        }




        public void Open(string path)
        {
            _excelApp.Workbooks.Open(path);
        }
        public void Close()
        {
            WorkBook.Close(0);
            //cleanup
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();



            //release com objects to fully kill excel process from running in the background
            if (XlRange != null)
                Marshal.ReleaseComObject(XlRange);
            Marshal.ReleaseComObject(XlWorksheet);

            //close and release

            Marshal.ReleaseComObject(WorkBook);

            //quit and release
            Application.Quit();
            _excelApp.Quit();
            Marshal.ReleaseComObject(_excelApp);
        }

    }

}
