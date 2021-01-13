using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCSV
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("please add srouce excel and destination csv file like (ExcelToCSV test.excel test.csv)");
                return -1;
            }
            try
            {
                var filePath = args[0];
                FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

                //Choose one of either 1 or 2
                IExcelDataReader excelReader;
                if (Path.GetExtension(filePath).Equals(".xls", StringComparison.CurrentCultureIgnoreCase))
                {
                    //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                excelReader.IsFirstRowAsColumnNames = true;
                DataSet result = excelReader.AsDataSet();
                var datatable = result.Tables[0];
                ToCSV(datatable, args[1]);
                Console.WriteLine("Make {0} successfully", args[1]);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return 0;
        }

        private static void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers    
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = string.Format("\"{0}\"", value.Replace(',', ' '));
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();


        }
    }
}
