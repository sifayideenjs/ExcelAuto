using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EC_to_SAP_Analysis_Console
{
    public class Converter
    {
        string _iCIMSFilePath = string.Empty;
        string _onbFilePath = string.Empty;
        string _sapHireFilePath = string.Empty;
        string _sapNameValidationFilePath = string.Empty;
        string _aliasCreationFilePath;

        private static object globalObj = new object();

        public Converter(string iCIMSFilePath, string onbFilePath, string sapHireFilePath, string sapNameValidationFilePath, string aliasCreationFilePath)
        {
            _iCIMSFilePath = iCIMSFilePath;
            _onbFilePath = onbFilePath;
            _sapHireFilePath = sapHireFilePath;
            _sapNameValidationFilePath = sapNameValidationFilePath;
            _aliasCreationFilePath = aliasCreationFilePath;
        }

        public bool Convert(string outputDirectory)
        {
            bool result = false;
            try
            {
                string fileName = ConfigurationManager.AppSettings["OutputFileName"];
                if (string.IsNullOrEmpty(fileName)) fileName = "EC to SAP Analysis";
                string saveFilePath = Path.Combine(outputDirectory, fileName + ".xlsx");

                IEnumerable<iCIMS> iCIMSs = null;
                try
                {
                    iCIMSs = ParseICIMS();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unable to parse " + "iCIMS" + " file");
                    Console.WriteLine(e.Message);
                    return result;
                }

                IEnumerable<ONB> ONBs = null;
                try
                {
                    ONBs = ParseONB();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unable to parse " + "ONB" + " file");
                    Console.WriteLine(e.Message);
                    return result;
                }

                IEnumerable<SAPNameValidation> SAPNameValidations = null;
                try
                {
                    SAPNameValidations = ParseSAPNameValidation();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unable to parse " + "SAP Name Validation" + " file");
                    Console.WriteLine(e.Message);
                    return result;
                }

                IEnumerable<SAPHireCompletionDateData> SAPHireCompletionDateDatas = null;
                try
                {
                    SAPHireCompletionDateDatas = ParseSAPHireCompletionDateData();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unable to parse " + "SAP Hire Completion Date Data" + " file");
                    Console.WriteLine(e.Message);
                    return result;
                }

                IEnumerable<AliasCreation> AliasCreations = null;
                try
                {
                    AliasCreations = ParseAliasCreation();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unable to parse " + "Alias Creation" + " file");
                    Console.WriteLine(e.Message);
                    return result;
                }

                List<EC_TO_SAP> EC_TO_SAPs = new List<EC_TO_SAP>();

                var iCIMSList = iCIMSs.Where(i => !string.IsNullOrEmpty(i.FullName));
                foreach (var iCIMS in iCIMSList)
                {
                    string FullName = iCIMS.FullName;
                    string StartDate = iCIMS.StartDate;
                    string PositionNumber = iCIMS.PositionNumber;
                    string RequistionID = iCIMS.RequistionID;
                    string Country = iCIMS.Country;
                    string LastPreOnboard = GetLongDateString(iCIMS.LastPreOnboard);
                    string EventReason = iCIMS.EventReason;

                    string PersonnelNumber = string.Empty;
                    string EmployeeName = string.Empty;
                    var SAPNameValidation = SAPNameValidations.Where(i => !string.IsNullOrEmpty(i.EmployeeName)).FirstOrDefault(i => i.EmployeeName.Trim() == FullName.Trim());
                    if (SAPNameValidation == null)
                    {
                        //Console.WriteLine("SAP Name Validation : " + FullName.Trim());
                        //continue;
                    }
                    else
                    {
                        PersonnelNumber = SAPNameValidation.PersonnelNumber;
                        //PersonnelNumber = PersonnelNumber.TrimStart('0');
                        EmployeeName = SAPNameValidation.EmployeeName;
                    }

                    string PostHireVerificationStepStartdate = string.Empty;
                    string PostHireVerificationStepEnddate = string.Empty;
                    string GlobalPreOnboardingStartdate = string.Empty;
                    string GlobalPreOnboardingEnddate = string.Empty;
                    var ONB = ONBs.Where(i => !string.IsNullOrEmpty(i.EmployeeLogin)).FirstOrDefault(i => i.EmployeeLogin.Trim() == PersonnelNumber.TrimStart('0'));
                    if (ONB == null)
                    {
                        //Console.WriteLine("ONB : " + FullName.Trim() + " - " + PersonnelNumber.Trim());
                        //continue;
                    }
                    else
                    {
                        PostHireVerificationStepStartdate = GetLongDateString(ONB.PostHireVerificationStepStartdate);                      
                        PostHireVerificationStepEnddate = GetLongDateString(ONB.PostHireVerificationStepEnddate);
                        GlobalPreOnboardingStartdate = GetLongDateString(ONB.GlobalPreOnboardingStartdate);
                        GlobalPreOnboardingEnddate = GetLongDateString(ONB.GlobalPreOnboardingEnddate);
                    }

                    string HireCompletedDateinSAP = string.Empty;
                    var SAPHireCompletionDateData = SAPHireCompletionDateDatas.Where(i => !string.IsNullOrEmpty(i.PersonnelNumber)).FirstOrDefault(i => i.PersonnelNumber.Trim() == PersonnelNumber.Trim());
                    if (SAPHireCompletionDateData == null)
                    {
                        //Console.WriteLine("SAP Hire Completion Date Data : " + FullName.Trim() + " - " + PersonnelNumber.Trim());
                        //continue;
                    }
                    else
                    {
                        HireCompletedDateinSAP = GetLongDateString(SAPHireCompletionDateData.HireCompletedDateinSAP);
                    }

                    string AliasCreationDate = string.Empty;
                    var AliasCreation = AliasCreations.Where(i => !string.IsNullOrEmpty(i.PersonnelNumber)).FirstOrDefault(i => i.PersonnelNumber.Trim() == PersonnelNumber);
                    if (AliasCreation == null)
                    {
                        //Console.WriteLine("Alias Creation : " + FullName.Trim() + " - " + PersonnelNumber.Trim());
                        //continue;
                    }
                    else
                    {
                        AliasCreationDate = GetLongDateString(AliasCreation.AliasCreationDate);
                    }

                    EC_TO_SAP ec_to_sap = new EC_TO_SAP()
                    {
                        FullName = FullName,
                        StartDate = StartDate,
                        PositionNumber = PositionNumber,
                        RequistionID = RequistionID,
                        Country = Country,
                        LastPreOnboard = LastPreOnboard,
                        PostHireVerificationStepStartdate = PostHireVerificationStepStartdate,
                        PostHireVerificationStepEnddate = PostHireVerificationStepEnddate,
                        GlobalPreOnboardingStartdate = GlobalPreOnboardingStartdate,
                        GlobalPreOnboardingEnddate = GlobalPreOnboardingEnddate,
                        PersonnelNumber = PersonnelNumber,
                        EmployeeName = EmployeeName,
                        EventReason = EventReason,
                        HireCompletedDateinSAP = HireCompletedDateinSAP,
                        AliasCreationDate = AliasCreationDate
                    };

                    EC_TO_SAPs.Add(ec_to_sap);
                }

                if(EC_TO_SAPs.Count > 0)
                {
                    DataTable dataTable = ToDataTable<EC_TO_SAP>(EC_TO_SAPs);
                    ExportTemplate(saveFilePath);

                    if (File.Exists(saveFilePath))
                    {
                        GenerateExcel(saveFilePath, dataTable);
                        Process.Start(saveFilePath);
                    }
                    else
                    {
                        Console.WriteLine("Output file alrady exist. Please delete it.");
                    }

                    result = true;
                }
                else
                {
                    Console.WriteLine("None of the data get matched with one another amount the input files. Please check it.");
                    result = false;
                }                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                result = false;
            }

            return result;
        }

        private string GetLongDateString(string dateString)
        {
            string longDateString = dateString;
            if (!string.IsNullOrEmpty(longDateString))
            {
                DateTime dt = DateTime.Parse(longDateString);
                longDateString = dt.ToShortDateString();
            }

            return longDateString;
        }

        private IEnumerable<iCIMS> ParseICIMS()
        {
            var excel = new ExcelQueryFactory(_iCIMSFilePath);
            excel.AddMapping("FullName", "Recruiting Workflow Profile (Person Full Name: First, Last Label");
            excel.AddMapping("StartDate", "Person : Start Date");
            excel.AddMapping("PositionNumber", "Requisition : Position no");
            excel.AddMapping("RequistionID", "Requisition : Req ID");
            excel.AddMapping("Country", "Requisition : Company Country");
            excel.AddMapping("LastPreOnboard", "Last Pre-Onboard: Hire Complete");
            excel.AddMapping("EventReason", "Event Reason");

            var iCIMSs = excel.Worksheet<iCIMS>(0).Select(p => p).ToList();
            return iCIMSs;
        }

        private IEnumerable<ONB> ParseONB()
        {
            var excel = new ExcelQueryFactory(_onbFilePath);
            excel.AddMapping("EmployeeLogin", "Employee Login");
            excel.AddMapping("PostHireVerificationStepStartdate", "Post Hire Verification Step Start Date");
            excel.AddMapping("PostHireVerificationStepEnddate", "Post Hire Verification Step End Date");
            excel.AddMapping("GlobalPreOnboardingStartdate", "Global Pre-Onboarding Start Date");
            excel.AddMapping("GlobalPreOnboardingEnddate", "Global Pre-Onboarding End Date");

            var ONBs = excel.Worksheet<ONB>(0).Select(p => p).ToList();
            return ONBs;
        }

        private IEnumerable<SAPNameValidation> ParseSAPNameValidation()
        {
            var excel = new ExcelQueryFactory(_sapNameValidationFilePath);
            excel.AddMapping("PersonnelNumber", "Personnel Number");
            excel.AddMapping("EmployeeName", "Employee Name");

            var SAPNameValidations = excel.Worksheet<SAPNameValidation>(0).Select(p => p).ToList();
            return SAPNameValidations;
        }

        private IEnumerable<SAPHireCompletionDateData> ParseSAPHireCompletionDateData()
        {
            IEnumerable<SAPHireCompletionDateData> SAPHireCompletionDateDatas = null;
            var excel = new ExcelQueryFactory(_sapHireFilePath);
            excel.AddMapping("PersonnelNumber", "PersNo");
            excel.AddMapping("HireCompletedDateinSAP", "Chngd on");

            SAPHireCompletionDateDatas = excel.Worksheet<SAPHireCompletionDateData>(0).Select(p => p).ToList();
            return SAPHireCompletionDateDatas;
        }

        private IEnumerable<AliasCreation> ParseAliasCreation()
        {
            var excel = new ExcelQueryFactory(_aliasCreationFilePath);
            excel.AddMapping("PersonnelNumber", "PersNo");
            excel.AddMapping("AliasCreationDate", "Chngd on");

            var AliasCreations = excel.Worksheet<AliasCreation>(0).Select(p => p).ToList();
            return AliasCreations;
        }

        private void ExportTemplate(string saveFilePath)
        {
            try
            {
                Assembly executingAssembly = Assembly.GetExecutingAssembly();
                string[] manifestResourceNames = executingAssembly.GetManifestResourceNames();
                foreach (string manifestResourceName in manifestResourceNames)
                {
                    if (manifestResourceName.StartsWith("EC_to_SAP_Analysis_Console.template.xlsx"))
                    {
                        lock (globalObj)
                        {
                            if (File.Exists(saveFilePath))
                            {
                                try
                                {
                                    File.Delete(saveFilePath);
                                }
                                catch(Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                            }

                            if (!File.Exists(saveFilePath))
                            {
                                using (var stream = executingAssembly.GetManifestResourceStream(manifestResourceName))
                                {
                                    using (FileStream fileStream = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                                    {
                                        int num;
                                        byte[] buffer = new byte[0x10000];
                                        while ((num = stream.Read(buffer, 0, buffer.Length)) > 0)
                                        {
                                            fileStream.Write(buffer, 0, num);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Output file alrady exist. Please delete it.");
                            }
                        }

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                var attribute = prop.GetCustomAttribute<DisplayNameAttribute>();
                if(attribute != null)
                {
                    dataTable.Columns.Add(attribute.DisplayName);
                }
                else
                {
                    dataTable.Columns.Add(prop.Name);
                }
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        private static void GenerateExcel(string excelFilePath, DataTable dataTable)
        {
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(excelFilePath, true))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                Sheet sheet = sheets.SingleOrDefault(s => s.Name == "EC 2 SAP Analysis");
                if (sheet == null) return;

                string relationshipId = sheet.Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);

                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                //DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in dataTable.Columns)
                {
                    columns.Add(column.ColumnName);

                    //DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    //cell.DataType = CellValues.String;
                    //cell.CellValue = new CellValue(column.ColumnName);
                    //headerRow.AppendChild(cell);
                }

                //sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in dataTable.Rows)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (String col in columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }
    }
}
