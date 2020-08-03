using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.AspNetCore.Http;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;

namespace TalkUnderSun
{
    public class CreateExcelFile
    {
        private const int DATE_FORMAT_ID = 1;

        public static bool CreateExcelDocument<T>(List<T> list, string xlsxFilePath)
        {
            return CreateExcelFile.CreateExcelDocument(new DataSet()
            {
                Tables =
                {
                    CreateExcelFile.ListToDataTable<T>(list)
                }
            }, xlsxFilePath);
        }

        /// <summary>
        /// List转换为DataTable
        /// </summary>
        /// <typeparam name="T">列表中的元素的类型</typeparam>
        /// <param name="list">列表元素</param>
        /// <returns></returns>
        public static DataTable ListToDataTable<T>(List<T> list)
        {
            DataTable dataTable = new DataTable();
            foreach (PropertyInfo propertyInfo in typeof (T).GetProperties())
                dataTable.Columns.Add(new DataColumn(propertyInfo.Name, CreateExcelFile.GetNullableType(propertyInfo.PropertyType)));
            foreach (T obj in list)
            {
                DataRow row = dataTable.NewRow();
                foreach (PropertyInfo propertyInfo in typeof (T).GetProperties())
                {
                    if (!CreateExcelFile.IsNullableType(propertyInfo.PropertyType))
                        row[propertyInfo.Name] = propertyInfo.GetValue((object) obj, (object[]) null);
                    else
                        row[propertyInfo.Name] = propertyInfo.GetValue((object) obj, (object[]) null) ?? (object) DBNull.Value;
                }
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }

        public static bool CreateExcelDocument(DataTable dt, string filename, HttpResponse response)
        {
            try
            {
                DataSet ds = new DataSet();
                ds.Tables.Add(dt);
                CreateExcelFile.CreateExcelDocumentAsStream(ds, filename, response, string.Empty);
                ds.Tables.Remove(dt);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }

        public static bool CreateExcelDocument<T>(List<T> list, string filename, HttpResponse response, string reportType)
        {
            try
            {
                CreateExcelFile.CreateExcelDocumentAsStream(new DataSet()
                {
                    Tables =
                    {
                        CreateExcelFile.ListToDataTable<T>(list)
                    }
                }, filename, response, reportType);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }

        public static bool CreateExcelDocumentAsStream(DataSet ds, string filename, HttpResponse response, string reportType)
        {
            try
            {
                MemoryStream memoryStream = new MemoryStream();
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create((Stream) memoryStream, SpreadsheetDocumentType.Workbook, true))
                    CreateExcelFile.WriteExcelFile(ds, spreadsheet, reportType);
                memoryStream.Flush();
                memoryStream.Position = 0L;
                response.ClearContent();
                response.Clear();
                response.Buffer = true;
                response.Charset = "";
                response.Cache.SetCacheability(HttpCacheability.NoCache);
                response.AddHeader("content-disposition", "attachment; filename=" + filename);
                response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                response.AddHeader("Content-Length", memoryStream.Length.ToString());
                byte[] buffer = new byte[memoryStream.Length];
                memoryStream.Read(buffer, 0, buffer.Length);
                memoryStream.Close();
                response.BinaryWrite(buffer);
                response.Flush();
                response.End();
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }

        public static bool CreateExcelDocument(DataSet ds, string excelFilename)
        {
            try
            {
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(excelFilename, SpreadsheetDocumentType.Workbook))
                    CreateExcelFile.WriteExcelFile(ds, spreadsheet, string.Empty);
                Trace.WriteLine("Successfully created: " + excelFilename);
                return true;
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Failed, exception thrown: " + ex.Message);
                return false;
            }
        }

        public static void WriteExcelFile(DataSet ds, SpreadsheetDocument spreadsheet, string reportType)
        {
            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new Workbook();
            spreadsheet.WorkbookPart.Workbook.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) new BookViews(new OpenXmlElement[1]
                {
                    (OpenXmlElement) new WorkbookView()
                })
            });
            spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles").Stylesheet = (Stylesheet) new CustomStylesheet();
            uint num1 = 1U;
            Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            foreach (DataTable dt in (InternalDataCollectionBase) ds.Tables)
            {
                string tableName = dt.TableName;
                WorksheetPart worksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                Sheet sheet = new Sheet()
                {
                    Id = (StringValue) spreadsheet.WorkbookPart.GetIdOfPart((OpenXmlPart) worksheetPart),
                    SheetId = (UInt32Value) num1,
                    Name = (StringValue) tableName
                };
                sheets.Append(new OpenXmlElement[1]
                {
                    (OpenXmlElement) sheet
                });
                CreateExcelFile.WriteDataTableToExcelWorksheet(dt, worksheetPart, reportType);
                ++num1;
            }
            WorksheetPart worksheetPart1 = Enumerable.First<WorksheetPart>(spreadsheet.WorkbookPart.WorksheetParts);
            SheetData sheetData = Enumerable.First<SheetData>(worksheetPart1.Worksheet.Elements<SheetData>());
            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.WorkbookStylesPart;
            Font font = new Font(new OpenXmlElement[3]
            {
                (OpenXmlElement) new Bold(),
                (OpenXmlElement) new FontSize()
                {
                    Val = (DoubleValue) 11.0
                },
                (OpenXmlElement) new FontName()
                {
                    Val = (StringValue) "Calibri"
                }
            });
            workbookStylesPart.Stylesheet.Fonts.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) font
            });
            workbookStylesPart.Stylesheet.Save();
            UInt32Value uint32Value = (UInt32Value) System.Convert.ToUInt32(workbookStylesPart.Stylesheet.Fonts.ChildElements.Count - 1);
            CellFormat cellFormat = new CellFormat()
            {
                FontId = uint32Value,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                ApplyFont = (BooleanValue) true
            };
            workbookStylesPart.Stylesheet.CellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat
            });
            Row row = Enumerable.First<Row>(sheetData.Elements<Row>());
            int num2 = workbookStylesPart.Stylesheet.CellFormats.ChildElements.Count - 1;
            foreach (CellType cellType in row.Elements<Cell>())
            {
                cellType.StyleIndex = (UInt32Value) System.Convert.ToUInt32(num2);
                worksheetPart1.Worksheet.Save();
            }
            spreadsheet.WorkbookPart.Workbook.Save();
        }

        private static void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart, string reportType)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create((OpenXmlPart) worksheetPart, Encoding.ASCII);
            writer.WriteStartElement((OpenXmlElement) new Worksheet());
            writer.WriteStartElement((OpenXmlElement) new SheetData());
            int count = dt.Columns.Count;
            bool[] flagArray1 = new bool[count];
            bool[] flagArray2 = new bool[count];
            string[] strArray = new string[count];
            for (int columnIndex = 0; columnIndex < count; ++columnIndex)
                strArray[columnIndex] = CreateExcelFile.GetExcelColumnName(columnIndex);
            uint num = 1U;
            writer.WriteStartElement((OpenXmlElement) new Row()
            {
                RowIndex = (UInt32Value) num
            });
            for (int index = 0; index < count; ++index)
            {
                DataColumn dataColumn = dt.Columns[index];
                CreateExcelFile.AppendTextCell(strArray[index] + "1", dataColumn.ColumnName, ref writer);
                flagArray1[index] = dataColumn.DataType.FullName == "System.Decimal" || dataColumn.DataType.FullName == "System.Int32" || dataColumn.DataType.FullName == "System.Double" || dataColumn.DataType.FullName == "System.Single";
                flagArray2[index] = dataColumn.DataType.FullName == "System.DateTime";
            }
            writer.WriteEndElement();
            foreach (DataRow dataRow in (InternalDataCollectionBase) dt.Rows)
            {
                ++num;
                writer.WriteStartElement((OpenXmlElement) new Row()
                {
                    RowIndex = (UInt32Value) num
                });
                for (int index = 0; index < count; ++index)
                {
                    string str = dataRow.ItemArray[index].ToString();
                    if (flagArray1[index])
                    {
                        double result = 0.0;
                        if (double.TryParse(str, out result))
                        {
                            string cellStringValue = result.ToString();
                            switch (reportType)
                            {
                                case "Commision":
                                    if (index == 12 || index == 17 || index == 23 || index == 24)
                                    {
                                        CreateExcelFile.AppendNumericCell(strArray[index] + num.ToString(), "$" + cellStringValue, ref writer);
                                        break;
                                    }
                                    CreateExcelFile.AppendNumericCell(strArray[index] + num.ToString(), cellStringValue, ref writer);
                                    break;
                                case "Broker":
                                    if (index == 4 || index == 7 || index == 12 || index == 14)
                                    {
                                        CreateExcelFile.AppendNumericCell(strArray[index] + num.ToString(), "$" + cellStringValue, ref writer);
                                        break;
                                    }
                                    CreateExcelFile.AppendNumericCell(strArray[index] + num.ToString(), cellStringValue, ref writer);
                                    break;
                                default:
                                    CreateExcelFile.AppendNumericCell(strArray[index] + num.ToString(), cellStringValue, ref writer);
                                    break;
                            }
                        }
                    }
                    else if (flagArray2[index])
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            string cellStringValue = DateTime.Parse(str).ToString("yyyy-MM-dd HH:mm:ss");
                            CreateExcelFile.AppendTextCell(strArray[index] + num.ToString(), cellStringValue, ref writer);
                        }
                        else
                            CreateExcelFile.AppendTextCell(strArray[index] + num.ToString(), str, ref writer);
                    }
                    else
                        CreateExcelFile.AppendTextCell(strArray[index] + num.ToString(), str, ref writer);
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.Close();
        }

        private static void AppendTextCell(string cellReference, string cellStringValue, ref OpenXmlWriter writer)
        {
            OpenXmlWriter openXmlWriter = writer;
            Cell cell1 = new Cell();
            cell1.CellValue = new CellValue(cellStringValue);
            cell1.CellReference = (StringValue) cellReference;
            cell1.DataType = (EnumValue<CellValues>) CellValues.String;
            Cell cell2 = cell1;
            openXmlWriter.WriteElement((OpenXmlElement) cell2);
        }

        private static void AppendDateCell(string cellReference, string cellStringValue, ref OpenXmlWriter writer)
        {
            OpenXmlWriter openXmlWriter = writer;
            Cell cell1 = new Cell();
            cell1.CellValue = new CellValue(cellStringValue);
            cell1.CellReference = (StringValue) cellReference;
            cell1.DataType = new EnumValue<CellValues>(CellValues.Number);
            cell1.StyleIndex = (UInt32Value) 0U;
            Cell cell2 = cell1;
            openXmlWriter.WriteElement((OpenXmlElement) cell2);
        }

        private static void AppendNumericCell(string cellReference, string cellStringValue, ref OpenXmlWriter writer)
        {
            OpenXmlWriter openXmlWriter = writer;
            Cell cell1 = new Cell();
            cell1.CellValue = new CellValue(cellStringValue);
            cell1.CellReference = (StringValue) cellReference;
            cell1.DataType = (EnumValue<CellValues>) CellValues.Number;
            Cell cell2 = cell1;
            openXmlWriter.WriteElement((OpenXmlElement) cell2);
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            if (columnIndex < 26)
                return ((char) (65 + columnIndex)).ToString();
            return string.Format("{0}{1}", (object) (char) (65 + columnIndex/26 - 1), (object) (char) (65 + columnIndex%26));
        }

        public static void SetSpreadsheetHeaderBold(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
            {
                WorksheetPart worksheetPart = Enumerable.First<WorksheetPart>(spreadsheetDocument.WorkbookPart.WorksheetParts);
                SheetData sheetData = Enumerable.First<SheetData>(worksheetPart.Worksheet.Elements<SheetData>());
                WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.WorkbookStylesPart;
                Font font = new Font(new OpenXmlElement[3]
                {
                    (OpenXmlElement) new Bold(),
                    (OpenXmlElement) new FontSize()
                    {
                        Val = (DoubleValue) 11.0
                    },
                    (OpenXmlElement) new FontName()
                    {
                        Val = (StringValue) "Calibri"
                    }
                });
                workbookStylesPart.Stylesheet.Fonts.Append(new OpenXmlElement[1]
                {
                    (OpenXmlElement) font
                });
                workbookStylesPart.Stylesheet.Save();
                UInt32Value uint32Value = (UInt32Value) System.Convert.ToUInt32(workbookStylesPart.Stylesheet.Fonts.ChildElements.Count - 1);
                CellFormat cellFormat = new CellFormat()
                {
                    FontId = uint32Value,
                    FillId = (UInt32Value) 0U,
                    BorderId = (UInt32Value) 0U,
                    ApplyFont = (BooleanValue) true
                };
                workbookStylesPart.Stylesheet.CellFormats.Append(new OpenXmlElement[1]
                {
                    (OpenXmlElement) cellFormat
                });
                Row row = Enumerable.First<Row>(sheetData.Elements<Row>());
                int num = workbookStylesPart.Stylesheet.CellFormats.ChildElements.Count - 1;
                foreach (CellType cellType in row.Elements<Cell>())
                {
                    cellType.StyleIndex = (UInt32Value) System.Convert.ToUInt32(num);
                    worksheetPart.Worksheet.Save();
                }
                spreadsheetDocument.Close();
            }
        }

        private static bool IsNullableType(Type type)
        {
            return type == typeof (string) || type.IsArray || type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof (Nullable<>));
        }

        public static bool CreateExcelDocument(DataTable dt, string excelFilename)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            bool excelDocument = CreateExcelFile.CreateExcelDocument(ds, excelFilename);
            ds.Tables.Remove(dt);
            return excelDocument;
        }

        private static Type GetNullableType(Type t)
        {
            Type type = t;
            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof (Nullable<>)))
                type = Nullable.GetUnderlyingType(t);
            return type;
        }
    }

    internal class CustomStylesheet : Stylesheet
    {
        public CustomStylesheet()
        {
            Fonts fonts = new Fonts();
            DocumentFormat.OpenXml.Spreadsheet.Font font1 = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName fontName1 = new FontName()
            {
                Val = StringValue.FromString("Arial")
            };
            FontSize fontSize1 = new FontSize()
            {
                Val = DoubleValue.FromDouble(11.0)
            };
            font1.FontName = fontName1;
            font1.FontSize = fontSize1;
            fonts.Append(new OpenXmlElement[1]
            {
                font1
            });
            DocumentFormat.OpenXml.Spreadsheet.Font font2 = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName fontName2 = new FontName()
            {
                Val = StringValue.FromString("Arial")
            };
            FontSize fontSize2 = new FontSize()
            {
                Val = DoubleValue.FromDouble(12.0)
            };
            font2.FontName = fontName2;
            font2.FontSize = fontSize2;
            font2.Bold = new Bold();
            fonts.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) font2
            });
            fonts.Count = UInt32Value.FromUInt32((uint) fonts.ChildElements.Count);
            Fills fills = new Fills();
            fills.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = (EnumValue<PatternValues>) PatternValues.None
                    }
                }
            });
            fills.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = (EnumValue<PatternValues>) PatternValues.Gray125
                    }
                }
            });
            fills.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = (EnumValue<PatternValues>) PatternValues.Solid
                    }
                }
            });
            fills.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) new Fill()
                {
                    PatternFill = new PatternFill()
                    {
                        PatternType = (EnumValue<PatternValues>) PatternValues.Solid
                    }
                }
            });
            fills.Count = UInt32Value.FromUInt32((uint) fills.ChildElements.Count);
            Borders borders = new Borders();
            Border border1 = new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) border1
            });
            Border border2 = new Border();
            Border border3 = border2;
            LeftBorder leftBorder1 = new LeftBorder();
            leftBorder1.Style = (EnumValue<BorderStyleValues>) BorderStyleValues.Thin;
            LeftBorder leftBorder2 = leftBorder1;
            border3.LeftBorder = leftBorder2;
            Border border4 = border2;
            RightBorder rightBorder1 = new RightBorder();
            rightBorder1.Style = (EnumValue<BorderStyleValues>) BorderStyleValues.Thin;
            RightBorder rightBorder2 = rightBorder1;
            border4.RightBorder = rightBorder2;
            Border border5 = border2;
            TopBorder topBorder1 = new TopBorder();
            topBorder1.Style = (EnumValue<BorderStyleValues>) BorderStyleValues.Thin;
            TopBorder topBorder2 = topBorder1;
            border5.TopBorder = topBorder2;
            Border border6 = border2;
            BottomBorder bottomBorder1 = new BottomBorder();
            bottomBorder1.Style = (EnumValue<BorderStyleValues>) BorderStyleValues.Thin;
            BottomBorder bottomBorder2 = bottomBorder1;
            border6.BottomBorder = bottomBorder2;
            border2.DiagonalBorder = new DiagonalBorder();
            Border border7 = border2;
            borders.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) border7
            });
            Border border8 = new Border();
            border8.LeftBorder = new LeftBorder();
            border8.RightBorder = new RightBorder();
            Border border9 = border8;
            TopBorder topBorder3 = new TopBorder();
            topBorder3.Style = (EnumValue<BorderStyleValues>) BorderStyleValues.Thin;
            TopBorder topBorder4 = topBorder3;
            border9.TopBorder = topBorder4;
            Border border10 = border8;
            BottomBorder bottomBorder3 = new BottomBorder();
            bottomBorder3.Style = (EnumValue<BorderStyleValues>) BorderStyleValues.Thin;
            BottomBorder bottomBorder4 = bottomBorder3;
            border10.BottomBorder = bottomBorder4;
            border8.DiagonalBorder = new DiagonalBorder();
            Border border11 = border8;
            borders.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) border11
            });
            borders.Count = UInt32Value.FromUInt32((uint) borders.ChildElements.Count);
            CellStyleFormats cellStyleFormats = new CellStyleFormats();
            CellFormat cellFormat1 = new CellFormat()
            {
                NumberFormatId = (UInt32Value) 0U,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U
            };
            cellStyleFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat1
            });
            cellStyleFormats.Count = UInt32Value.FromUInt32((uint) cellStyleFormats.ChildElements.Count);
            uint num1 = 164U;
            NumberingFormats numberingFormats = new NumberingFormats();
            CellFormats cellFormats = new CellFormats();
            CellFormat cellFormat2 = new CellFormat()
            {
                NumberFormatId = (UInt32Value) 0U,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat2
            });
            NumberingFormat numberingFormat1 = new NumberingFormat();
            NumberingFormat numberingFormat2 = numberingFormat1;
            int num2 = (int) num1;
            int num3 = 1;
            uint num4 = (uint) (num2 + num3);
            UInt32Value uint32Value1 = UInt32Value.FromUInt32((uint) num2);
            numberingFormat2.NumberFormatId = uint32Value1;
            numberingFormat1.FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss");
            NumberingFormat numberingFormat3 = numberingFormat1;
            numberingFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) numberingFormat3
            });
            NumberingFormat numberingFormat4 = new NumberingFormat();
            NumberingFormat numberingFormat5 = numberingFormat4;
            int num5 = (int) num4;
            int num6 = 1;
            uint num7 = (uint) (num5 + num6);
            UInt32Value uint32Value2 = UInt32Value.FromUInt32((uint) num5);
            numberingFormat5.NumberFormatId = uint32Value2;
            numberingFormat4.FormatCode = StringValue.FromString("#,##0.0000");
            NumberingFormat numberingFormat6 = numberingFormat4;
            numberingFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) numberingFormat6
            });
            NumberingFormat numberingFormat7 = new NumberingFormat();
            NumberingFormat numberingFormat8 = numberingFormat7;
            int num8 = (int) num7;
            int num9 = 1;
            uint num10 = (uint) (num8 + num9);
            UInt32Value uint32Value3 = UInt32Value.FromUInt32((uint) num8);
            numberingFormat8.NumberFormatId = uint32Value3;
            numberingFormat7.FormatCode = StringValue.FromString("#,##0.00");
            NumberingFormat numberingFormat9 = numberingFormat7;
            numberingFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) numberingFormat9
            });
            NumberingFormat numberingFormat10 = new NumberingFormat()
            {
                NumberFormatId = UInt32Value.FromUInt32(num10),
                FormatCode = StringValue.FromString("@")
            };
            numberingFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) numberingFormat10
            });
            CellFormat cellFormat3 = new CellFormat()
            {
                NumberFormatId = (UInt32Value) 14U,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat3
            });
            CellFormat cellFormat4 = new CellFormat()
            {
                NumberFormatId = (UInt32Value) 4U,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat4
            });
            CellFormat cellFormat5 = new CellFormat()
            {
                NumberFormatId = numberingFormat3.NumberFormatId,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat5
            });
            CellFormat cellFormat6 = new CellFormat()
            {
                NumberFormatId = numberingFormat6.NumberFormatId,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat6
            });
            CellFormat cellFormat7 = new CellFormat()
            {
                NumberFormatId = numberingFormat9.NumberFormatId,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat7
            });
            CellFormat cellFormat8 = new CellFormat()
            {
                NumberFormatId = numberingFormat10.NumberFormatId,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat8
            });
            CellFormat cellFormat9 = new CellFormat()
            {
                NumberFormatId = numberingFormat10.NumberFormatId,
                FontId = (UInt32Value) 1U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 0U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat9
            });
            CellFormat cellFormat10 = new CellFormat()
            {
                NumberFormatId = numberingFormat10.NumberFormatId,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 0U,
                BorderId = (UInt32Value) 1U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat10
            });
            CellFormat cellFormat11 = new CellFormat()
            {
                NumberFormatId = numberingFormat9.NumberFormatId,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 2U,
                BorderId = (UInt32Value) 2U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat11
            });
            CellFormat cellFormat12 = new CellFormat()
            {
                NumberFormatId = numberingFormat10.NumberFormatId,
                FontId = (UInt32Value) 0U,
                FillId = (UInt32Value) 2U,
                BorderId = (UInt32Value) 2U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat12
            });
            CellFormat cellFormat13 = new CellFormat()
            {
                NumberFormatId = numberingFormat10.NumberFormatId,
                FontId = (UInt32Value) 1U,
                FillId = (UInt32Value) 3U,
                BorderId = (UInt32Value) 2U,
                FormatId = (UInt32Value) 0U,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormat13
            });
            numberingFormats.Count = UInt32Value.FromUInt32((uint) numberingFormats.ChildElements.Count);
            cellFormats.Count = UInt32Value.FromUInt32((uint) cellFormats.ChildElements.Count);
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) numberingFormats
            });
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) fonts
            });
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) fills
            });
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) borders
            });
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellStyleFormats
            });
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellFormats
            });
            CellStyles cellStyles = new CellStyles();
            CellStyle cellStyle = new CellStyle()
            {
                Name = StringValue.FromString("Normal"),
                FormatId = (UInt32Value) 0U,
                BuiltinId = (UInt32Value) 0U
            };
            cellStyles.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellStyle
            });
            cellStyles.Count = UInt32Value.FromUInt32((uint) cellStyles.ChildElements.Count);
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) cellStyles
            });
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) new DifferentialFormats()
                {
                    Count = (UInt32Value) 0U
                }
            });
            this.Append(new OpenXmlElement[1]
            {
                (OpenXmlElement) new TableStyles()
                {
                    Count = (UInt32Value) 0U,
                    DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                    DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
                }
            });
        }

        private static ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
        {
            ForegroundColor foregroundColor = new ForegroundColor();
            foregroundColor.Rgb = new HexBinaryValue()
            {
                Value = ColorTranslator.ToHtml(System.Drawing.Color.FromArgb((int) fillColor.A, (int) fillColor.R, (int) fillColor.G, (int) fillColor.B)).Replace("#", "")
            };
            return foregroundColor;
        }
    }
}

