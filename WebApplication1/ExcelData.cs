using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Web;

namespace WebApplication1
{
    public class ExcelData
    {
        string _path;
        public ExcelData(string path)
        {
            _path = path;
        }

        public IExcelDataReader GetExcelReader()
        {
            FileStream stream = File.Open(_path, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = null;
            try
            {
                if (_path.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                if (_path.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                return reader;
            }
            catch (Exception)
            {
                throw;
            }
        }

        //read the sheets name if you need
        public IEnumerable<string> GetWorksheetNames()
        {
            var reader = this.GetExcelReader();
            var workbook = reader.AsDataSet();
            var sheets = from DataTable sheet in workbook.Tables select sheet.TableName;
            return sheets;
        }

        //read data in a specified sheet
        public IEnumerable<DataRow> GetData(string sheet)
        {

            var reader = this.GetExcelReader();
            var workSheet = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    //indicates if use the header values
                    UseHeaderRow = true
                }

            }).Tables[sheet];

            var rows = from DataRow a in workSheet.Rows select a;
            return rows;
        }

    }
}