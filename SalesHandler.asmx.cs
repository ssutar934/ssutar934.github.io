using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Script.Services;
using System.Web.Services;

namespace Sales
{
    /// <summary>
    /// Summary description for SalesHandler
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    [System.Web.Script.Services.ScriptService]
    public class SalesHandler : System.Web.Services.WebService
    {

        [WebMethod]
        public string HelloWorld()
        {
            return "Hello World";
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public string GetOrderData()
        {
            string filePath = "OrderData\\OrderDetails.xlsx";
            filePath = Server.MapPath(filePath);
            int sheetNumber = 0;

            SpreadsheetGear.IWorkbookSet workbookSet = SpreadsheetGear.Factory.GetWorkbookSet();
            SpreadsheetGear.IWorkbook workbook = workbookSet.Workbooks.Open(filePath);
            //SpreadsheetGear.IWorkbook workbook = SpreadsheetGear.Factory.GetWorkbook(filePath);

            // Get a range from an existing defined name.
            SpreadsheetGear.IRange range = workbook.Worksheets[sheetNumber].UsedRange;

            SpreadsheetGear.Data.GetDataFlags flags = SpreadsheetGear.Data.GetDataFlags.FormattedText;

            // Create a new DataTable.
            DataTable dataTable = new DataTable();

            try
            {
                // Get a reference to the worksheet.
                SpreadsheetGear.IWorksheet worksheet = range.Worksheet;

                // Get a reference to all the worksheet cells.
                SpreadsheetGear.IRange cells = worksheet.Cells;

                // Get a reference to the advanced API.
                SpreadsheetGear.Advanced.Cells.IValues values =
                    (SpreadsheetGear.Advanced.Cells.IValues)worksheet;



                // Determine the row and column coordinates of the range.
                int row1 = range.Row;
                int col1 = range.Column;
                int rowCount = range.RowCount;
                int colCount = range.ColumnCount;
                int row2 = row1 + rowCount - 1;
                int col2 = col1 + colCount - 1;
                int row = row1;

                // If the first row is not used for column headers...
                if ((flags & SpreadsheetGear.Data.GetDataFlags.NoColumnHeaders) != 0)
                {
                    // Create columns using simple column names.
                    for (int col = col1; col <= col2; col++)
                    {
                        string colName = "Column" + (col - col1 + 1);
                        dataTable.Columns.Add(colName);
                    }
                }
                else
                {
                    // Create columns using the first row in the range for column names.
                    for (int col = col1; col <= col2; col++)
                    {
                        // Use the IRange API to get formatted text.
                        string colName = cells[row, col].Text;

                        dataTable.Columns.Add(colName);
                    }
                    row++;
                }

                // If the DataTable column data types should be set...
                if ((flags & SpreadsheetGear.Data.GetDataFlags.NoColumnTypes) == 0 && row <= row2)
                {
                    for (int col = col1; col <= col2; col++)
                    {
                        // Get a reference to the DataTable column.
                        System.Data.DataColumn dataCol = dataTable.Columns[col - col1];

                        // If formatted text is to be used for all cell values...
                        if ((flags & SpreadsheetGear.Data.GetDataFlags.FormattedText) != 0)
                        {
                            // Set the data type to a string.
                            dataCol.DataType = typeof(string);
                        }
                        else
                        {
                            // Set the data type based on the type of data in the cell.
                            //
                            // Note that this will cause problems if a column does not contain
                            // consistent data types - for example a column of formulas where
                            // the first is numeric but one of the following is an error.
                            SpreadsheetGear.Advanced.Cells.IValue value = values[row, col];
                            if (value != null)
                            {
                                switch (value.Type)
                                {
                                    case SpreadsheetGear.Advanced.Cells.ValueType.Number:
                                        dataCol.DataType = typeof(double);
                                        break;
                                    case SpreadsheetGear.Advanced.Cells.ValueType.Text:
                                    case SpreadsheetGear.Advanced.Cells.ValueType.Error:
                                        dataCol.DataType = typeof(string);
                                        break;
                                    case SpreadsheetGear.Advanced.Cells.ValueType.Logical:
                                        dataCol.DataType = typeof(bool);
                                        break;
                                }
                            }
                        }
                    }
                }

                // If formatted text is to be used for all cell values...
                if ((flags & SpreadsheetGear.Data.GetDataFlags.FormattedText) != 0)
                {
                    // Create the row data as an array of strings.
                    string[] rowData = new string[colCount];
                    for (; row <= row2; row++)
                    {
                        // If the row is not hidden...
                        if (!cells[row, 0].Rows.Hidden)
                        {
                            for (int col = col1; col <= col2; col++)
                            {
                                // Use the IRange API to get formatted text.
                                string text = cells[row, col].Text;
                                rowData[col - col1] = text;
                            }

                            // Add a new row using the array of formatted strings.
                            dataTable.Rows.Add(rowData);
                        }
                    }
                }
                else
                {
                    // Create the row data as an array of objects.
                    object[] rowData = new object[colCount];
                    for (; row <= row2; row++)
                    {
                        // If the row is not hidden...
                        if (!cells[row, 0].Rows.Hidden)
                        {
                            for (int col = col1; col <= col2; col++)
                            {
                                // Use the advanced API to get the raw data values.
                                SpreadsheetGear.Advanced.Cells.IValue value = values[row, col];
                                object obj = null;
                                if (value != null)
                                {
                                    switch (value.Type)
                                    {
                                        case SpreadsheetGear.Advanced.Cells.ValueType.Number:
                                            obj = value.Number;
                                            break;
                                        case SpreadsheetGear.Advanced.Cells.ValueType.Text:
                                            obj = value.Text;
                                            break;
                                        case SpreadsheetGear.Advanced.Cells.ValueType.Logical:
                                            obj = value.Logical;
                                            break;
                                        case SpreadsheetGear.Advanced.Cells.ValueType.Error:
                                            // This will create problems if it is a column type
                                            // of double or bool.
                                            obj = "#" + value.Error.ToString().ToUpper() + "!";
                                            break;
                                    }
                                }
                                rowData[col - col1] = obj;
                            }

                            // Add a new row using the array of objects.
                            dataTable.Rows.Add(rowData);
                        }
                    }
                }




                if (dataTable == null)
                {
                    return null;
                }

                System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();

                serializer.MaxJsonLength = 2147483644;

                List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                Dictionary<string, object> jrow;
                foreach (DataRow dr in dataTable.Rows)
                {
                    jrow = new Dictionary<string, object>();
                    foreach (DataColumn col in dataTable.Columns)
                    {
                        //Rename the column is they requested an alias
                        string colname = col.ColumnName;
                        if (!jrow.ContainsKey(colname))
                            jrow.Add(colname, Convert.ToString(dr[col]));
                    }
                    rows.Add(jrow);
                }

                string json = "{\"Records\":" + serializer.Serialize(rows) + "}";//serializer.Serialize(rows); 

                return json;

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return null;
        }
    }
}
