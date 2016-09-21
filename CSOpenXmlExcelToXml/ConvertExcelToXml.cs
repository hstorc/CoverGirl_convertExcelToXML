/****************************** Module Header ******************************\
* Module Name:  ConvertExcelToXml.cs
* Project:      CSOpenXmlExcelToXml
* Copyright(c)  Microsoft Corporation.
* 
* This class is used to convert excel data to XML format string using Open XMl
* Firstly, we use OpenXML API to get data from Excel to DataTable
* Then we Load the DataTable to Dataset.
* At Last,we call DataSet.GetXml() to get XML format string 
*
* This source is subject to the Microsoft Public License.
* See http://www.microsoft.com/en-us/openness/licenses.aspx.
* All other rights reserved.
* 
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
* EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
* WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/


using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using System.Xml.Linq;
namespace CSOpenXmlExcelToXml
{
    public class ConvertExcelToXml
    {
        private string _excelfilename;

        /// <summary>
        ///  Read Data from selected excel file into DataTable
        /// </summary>
        /// <param name="filename">Excel File Path</param>
        /// <returns></returns>
        /// 
        private DataTable ReadExcelFile(string filename,string tablename, string[] fieldstotrim )
        {
            // Initialize an instance of DataTable
            DataTable dt = new DataTable();

            dt.TableName = tablename;
            int rowcounter = 0;
            try
            {

                // Use SpreadSheetDocument class of Open XML SDK to open excel file
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, false))
                {
                    // Get Workbook Part of Spread Sheet Document
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                    // Get all sheets in spread sheet document 
                    IEnumerable<Sheet> sheetcollection = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                    // Get relationship Id
                    string relationshipId = sheetcollection.First().Id.Value;

                    // Get sheet1 Part of Spread Sheet Document
                    WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);

                    // Get Data in Excel file
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    IEnumerable<Row> rowcollection = sheetData.Descendants<Row>();

                    if (rowcollection.Count() == 0)
                    {
                        return dt;
                    }

                    // Add columns
                    foreach (Cell cell in rowcollection.ElementAt(0))
                    {
                        dt.Columns.Add(GetValueOfCell(spreadsheetDocument, cell));
                    }

                    // Add rows into DataTable
                    foreach (Row row in rowcollection)
                    {
                        DataRow temprow = dt.NewRow();
                        int columnIndex = 0;
                        string columnname = "";
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            // Get Cell Column Index
                            columnname = GetColumnName(cell.CellReference);
                            int cellColumnIndex = GetColumnIndex(columnname);

                            if (columnIndex < cellColumnIndex)
                            {
                                do
                                {
                                    try
                                    {
                                        temprow[columnIndex] = string.Empty;
                                        columnIndex++;
                                    }
                                    catch (Exception x)
                                    {
                                        System.Diagnostics.Debug.Print(x.Message);

                                    }
                                }

                                while (columnIndex < cellColumnIndex);
                            }
                            if (fieldstotrim!=null && fieldstotrim.Where(f => f.Equals(columnname)).Count() == 1)
                            {
                                temprow[columnIndex] = GetValueOfCell(spreadsheetDocument, cell).Replace(" ","");
                            }
                            else {
                                temprow[columnIndex] = GetValueOfCell(spreadsheetDocument, cell);
                            }
                            columnIndex++;
                            if (columnIndex >= 7)
                                break;
                        }

                        // Add the row to DataTable
                        // the rows include header row
                        dt.Rows.Add(temprow);

                        rowcounter++;
                    }
                }

                // Here remove header row
                dt.Rows.RemoveAt(0);
                return dt;
            }
            catch (IOException ex)
            {
                throw new IOException(ex.Message);
            }
        }

        /// <summary>
        ///  Get Value of Cell 
        /// </summary>
        /// <param name="spreadsheetdocument">SpreadSheet Document Object</param>
        /// <param name="cell">Cell Object</param>
        /// <returns>The Value in Cell</returns>
        private static string GetValueOfCell(SpreadsheetDocument spreadsheetdocument, Cell cell)
        {
            // Get value in Cell
            SharedStringTablePart sharedString = spreadsheetdocument.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null)
            {
                return string.Empty;
            }

            string cellValue = cell.CellValue.InnerText;

            // The condition that the Cell DataType is SharedString
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return sharedString.SharedStringTable.ChildElements[int.Parse(cellValue)].InnerText;
            }
            else
            {
                return cellValue;
            }
        }

        /// <summary>
        /// Get Column Name From given cell name
        /// </summary>
        /// <param name="cellReference">Cell Name(For example,A1)</param>
        /// <returns>Column Name(For example, A)</returns>
        private string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name of cell
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

        /// <summary>
        /// Get Index of Column from given column name
        /// </summary>
        /// <param name="columnName">Column Name(For Example,A or AA)</param>
        /// <returns>Column Index</returns>
        private int GetColumnIndex(string columnName)
        {
            int columnIndex = 0;
            int factor = 1;

            // From right to left
            for (int position = columnName.Length - 1; position >= 0; position--)
            {
                // For letters
                if (Char.IsLetter(columnName[position]))
                {
                    columnIndex += factor * ((columnName[position] - 'A') + 1) - 1;
                    factor *= 26;
                }
            }

            return columnIndex;
        }

        public string GetXML(string filename)
        {
            _excelfilename = filename;

            XmlDocument xdoc = new XmlDocument();
            List<MainProductFields> main = GetMainProducts();
            List<UPC> upcsList = GetUPCs();
            XmlNode products = xdoc.CreateNode(XmlNodeType.Element, "Products", "");
            //XmlNode product, productUniqueID, name, productUrl, imageUrl, description, UPCs;
            //string nodeName = "";
            int uniqueId = 1;
            var produclist = new XElement("Feed",
                new XAttribute("xmlnsxsi", "http://www.w3.org/2001/XMLSchema-instance"),
                new XElement("Products",
                from product in main
                select new XElement("Product",
                        new XElement("ProductUniqueID", uniqueId++),
                        new XElement("Name", product.Name),
                        new XElement("ProductUrl", product.ProductURL),
                        new XElement("ImageUrl", product.ImageURL),
                        new XElement("Description", product.Description),
                        new XElement("UPCs",
                            from upcvalue in upcsList
                            where upcvalue.ProductName.Equals(product.Name)
                            select new XElement("UPC", upcvalue.UPCValue),
                            from gtinvalue in upcsList
                            where gtinvalue.ProductName.Equals(product.Name)
                            select new XElement("UPC", gtinvalue.GTINValue)
                            )
                           )
                          ));

            return produclist.ToString();
            //return xdoc.ToString();
        }

        private List<UPC> GetUPCs()
        {
            DataSet ds = new DataSet("Feeds");
            DataTable dt = ReadExcelFile(_excelfilename,"UPCs",new string[]{ "G"});
            var mainTable = (from p in dt.AsEnumerable()
                             group p by new UPC
                             {
                                 ProductName = p.Field<string>("ProductName"),
                                 UPCValue = p.Field<string>("UPC"),
                                 GTINValue = p.Field<string>("GTIN")
                             }
            into grp
                             select new
                             {
                                 grp.Key
                             });
            List<UPC> main = new List<UPC>();
            main.AddRange(mainTable.Select(x => new UPC
            {
                ProductName = x.Key.ProductName,
                UPCValue = x.Key.UPCValue,
                GTINValue = x.Key.GTINValue
            }));
            /*List<UPC> onecolumnupc = new List<UPC>();
            foreach (UPC upcvalue in main)
            {
                onecolumnupc.Add(new UPC { ProductName = upcvalue.ProductName, UPCValue = upcvalue.UPCValue });
                onecolumnupc.Add(new UPC { ProductName = upcvalue.ProductName, UPCValue = upcvalue.GTINValue });
            }*/

            //return onecolumnupc;
            return main;
        }

        private List<MainProductFields> GetMainProducts()
        {
            DataTable dt = ReadExcelFile(this._excelfilename,"Products",null);
            var mainTable = (from p in dt.AsEnumerable()
                             group p by new
                             MainProductFields
                             {
                                 // ProductUniqueID = p.Field<string>("ProductUniqueID"),
                                 Name = p.Field<string>("ProductName"),
                                 ProductURL = p.Field<string>("ProductURL"),
                                 ImageURL = p.Field<string>("ImageURL"),
                                 Description = p.Field<string>("Description")
                             } into grp
                             select new
                             {
                                 grp.Key
                             });
            List<MainProductFields> main = new List<MainProductFields>();
            main.AddRange(mainTable.Select(x => new MainProductFields
            {
                ProductUniqueID = x.Key.ProductUniqueID,
                Name = x.Key.Name,
                ImageURL = x.Key.ImageURL,
                ProductURL = x.Key.ProductURL,
                Description = x.Key.Description
            }));
            //main has unique records

            return main;
        }

    }

    public struct MainProductFields
    {
        public string ProductUniqueID;
        public string Name;
        public string ProductURL;
        public string ImageURL;
        public string Description;

    }
    public struct UPC
    {
        public string ProductName;
        public string UPCValue;
        public string GTINValue;
    }
}
