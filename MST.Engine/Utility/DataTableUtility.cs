using EPPlus.Extensions;
using OfficeOpenXml;
using System.Data;
using System.IO;

namespace MST.Engine.Utility
{
    class DataTableUtility
    {
        public static DataTable GetDataTable()
        {
            string reportFile = InitializeParamter.ReportFile;
            if (File.Exists(reportFile))
            {
                var package = new ExcelPackage(new FileInfo(reportFile));
                DataSet ds = Extensions.ToDataSet(package, true);
                if (ds.Tables.Contains(InitializeParamter.RunDay))
                {                    
                    return ds.Tables[InitializeParamter.RunDay];
                }
                else
                {
                    DataTable table = new DataTable();
                    table.Columns.Add("Suite Name", typeof(string));
                    table.Columns.Add("Case Name", typeof(string));
                    table.Columns.Add("Status", typeof(string));
                    table.Columns.Add("Start Time", typeof(string));
                    table.Columns.Add("End Time", typeof(string));
                    table.Columns.Add("Elapsed Time", typeof(double));
                    table.Columns.Add("Error Message", typeof(string));
                    table.Columns.Add("Result Path", typeof(string));
                    return table;
                }                                  
            }
            else
            {
                // Here we create a DataTable with four columns.
                DataTable table = new DataTable();
                table.Columns.Add("Suite Name", typeof(string));
                table.Columns.Add("Case Name", typeof(string));
                table.Columns.Add("Status", typeof(string));
                table.Columns.Add("Start Time", typeof(string));
                table.Columns.Add("End Time", typeof(string));
                table.Columns.Add("Elapsed Time", typeof(double));
                table.Columns.Add("Error Message", typeof(string));
                table.Columns.Add("Result Path", typeof(string));
                return table;
            }
        }
    }
}
