using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using RCMSTaskScheduler.GlobalCode;
using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net.Mail;
using System.Reflection;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RCMSTaskScheduler
{
    class Program
    {
        public static string RM97_ConnectionStr = "";
        public static string ProdOrDevPath = "";
        public static DateTime DateToday = DateTime.Now;
        public static DataTable dtProfile = new DataTable();
        public static ArrayList AttachedFiles = new ArrayList();
        public static string LogFile = "";
        public static string ReportsFolder = "";
        public static string RCMSFolder = "";
        public static string ProfileName = "";
        public static string RunString = "";
        public static string ErrMsg = "";
        public static string[] args = null;

        [STAThread]
        static void Main()
        {
            try
            {
                if (Environment.GetCommandLineArgs().Length == 1)//No arguments passed...
                {
                    Application.Exit();
                }

                args = Environment.GetCommandLineArgs();

                //Arguments...
                //=============================================================================================================
                //args[1] = RunString. Currently runs profile in normal or test mode but can be used 
                //          to process any job we want...
                //
                //args[2] = ProfileID. Profile to be Processed.
                //args[3] = Tester's email address.
                //args[4] = TaskName to be deleted. (deletes task after run in Test Mode)

                RunString = args[1];

                SetConnStringAndFolders();
                MiscFunctions.LogMsg("___________________________________________________", LogFile);
                MiscFunctions.LogMsg("Process Started: " + DateTime.Now, LogFile);
                if (RunString == "UpdateGAL")//Update GAL table (Global Address List)...
                {
                    UpDateGAL();
                }

                dtProfile = MiscFunctions.GetProfile(args[2]);
                ProfileName = dtProfile.Rows[0].Field<string>("ProfileName");
                MiscFunctions.LogMsg("Profile ID: " + args[2], LogFile);

                if (RunString == "RunSchedulerProfileTest")//testing...
                {
                    MiscFunctions.LogMsg("Testing Profile Name: " + ProfileName, LogFile);
                    MiscFunctions.LogMsg("Reqester: " + args[3], LogFile);
                }
                else
                {
                    MiscFunctions.LogMsg("Profile Name: " + ProfileName, LogFile);
                    MiscFunctions.LogMsg("Requester: " + dtProfile.Rows[0].Field<string>("Requester"), LogFile);
                }

                if (RunProfile(ref ErrMsg))
                {
                    MiscFunctions.LogMsg("Processed successfully!", LogFile);
                    Application.Exit();
                }
                else
                {
                    MiscFunctions.LogMsg("Process failed! Error: " + ErrMsg, LogFile);
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("Index was outside"))
                {
                    MiscFunctions.LogMsg("Fatal Error: " + ex.Message, LogFile);
                }
                Application.Exit();
            }

        }

        public static void UpDateGAL()
        {
            try
            {
                string ErrMsg = "";
                ArrayList GalArray = new ArrayList();
                Outlook.Application OutlookApp = new Outlook.Application();
                Outlook.NameSpace nameSpace = OutlookApp.GetNamespace("MAPI");
                nameSpace.Logon("Outlook", Missing.Value, false, true);
                Outlook.AddressList addrList = OutlookApp.Session.GetGlobalAddressList();
                //populate GalArray...
                foreach (Outlook.AddressEntry addressEntry in addrList.AddressEntries)
                {
                    if (addressEntry.DisplayType == Outlook.OlDisplayType.olUser)
                    {
                        if (addressEntry.Type == "EX")
                        {
                            Outlook.ExchangeUser exchUser = addressEntry.GetExchangeUser();
                            string email = exchUser.PrimarySmtpAddress.ToLower();
                            if (!email.Contains("rcmsreports@prezero.us") &
                                !email.Contains(".onmicrosoft.com"))
                            {
                                string Name = exchUser.Name == null ? "" : exchUser.Name;
                                string Mobile = exchUser.MobileTelephoneNumber == null ? "" : exchUser.MobileTelephoneNumber;
                                string JobTitle = exchUser.JobTitle == null ? "" : exchUser.JobTitle;
                                string Department = exchUser.Department == null ? "" : exchUser.Department;
                                string Location = exchUser.OfficeLocation == null ? "" : exchUser.OfficeLocation;

                                GalArray.Add(email + "|"
                                      + Name + "|"
                                      + Mobile + "|"
                                      + JobTitle + "|"
                                      + Department + "|"
                                      + Location
                                      );
                            }
                        }
                    }
                }

                //If everything made it this far, Delete all records in GAL since array is populated...
                MiscFunctions.RunGenericSQLCmd("delete from rm97.dbo.GAL", ref ErrMsg);

                //insert GAL address info....
                foreach (string addStr in GalArray)
                {
                    string[] FieldsArray = addStr.Split('|');
                    MiscFunctions.RunGenericSQLCmd("insert rm97.dbo.GAL (EmailAddress,Name,Mobile,JobTitle,Department,Location) " +
                    "VALUES ('" + FieldsArray[0] + "','" +
                    FieldsArray[1] + "','" + FieldsArray[2] + "','" + FieldsArray[3] + "','" +
                    FieldsArray[4] + "','" + FieldsArray[5] + "')", ref ErrMsg);
                }

                DataTable dt = MiscFunctions.RunGenericSelProc("select lower(email) as email from rcms_web.dbo.AspNetUsers", ref ErrMsg);
                foreach (DataRow row in dt.Rows)
                {
                    if (!MiscFunctions.AddressExist(row.Field<string>("email")))
                    {
                        MiscFunctions.RunGenericSQLCmd("insert rm97.dbo.GAL (EmailAddress,Name,Mobile,JobTitle,Department,Location) " +
                        "VALUES ('" + row.Field<string>("email") + "','','','','','')", ref ErrMsg);
                    }
                }

                int CountAddresses = MiscFunctions.GetGalCount();
                MiscFunctions.LogMsg("Updated GAL successfully! Total addresses= " + CountAddresses, LogFile);
                nameSpace = null;
                Application.Exit();
            }
            catch (Exception ex)
            {
                MiscFunctions.LogMsg("Updated GAL Failed! Error= " + ex.Message, LogFile);
                Application.Exit();
            }
        }
        public static bool RunProfile(ref string ErrMsg)
        {
            //Run Report(s) based on Report Criteria Options in Profile: (DateRange, Parameter1, Parameter2, ExportType)
            //Export in correct format (Excel/Pdf), attach to email and send from rcmsreports@prezero.us to profile recipients.
            //Log process activity in Logfile...
            //======================================================================================================
            try
            {
                //Initialize...
                DataTable dtSchedulerReports = new DataTable();
                string DateRangeString = "";
                string DateRangeValue = "";
                string CrystalParameter1Name = "";
                string CrystalParameter1ValueType = "";
                string CrystalParameter1Value = "";
                string CrystalParameter2Name = "";
                string CrystalParameter2ValueType = "";
                string CrystalParameter2Value = "";
                string ExportType = "";
                DateTime StartDate = DateTime.MinValue;
                DateTime EndDate = DateTime.MinValue;
                string EmailPswd = MiscFunctions.GetEmailPassWord();

                //First Get Reports. If no reports, log error and exit...
                if (dtProfile.Rows[0].Field<string>("Reports") == "")
                {
                    ErrMsg = "No reports in profile.";
                    return false;
                }

                //Sping through reports to create export files to be attached to email...
                string[] ReportFileNames = dtProfile.Rows[0].Field<string>("Reports").Split(';');
                string FullReportFileName = "";
                foreach (string ReportFileName in ReportFileNames)
                {
                    string TrimmedReportFileName = ReportFileName.Trim();
                    TrimmedReportFileName = ReportFileName.Trim();
                    FullReportFileName = MiscFunctions.GetReportPathFromFileName(TrimmedReportFileName) + TrimmedReportFileName;
                    dtSchedulerReports = MiscFunctions.RunGenericSelProc("Select * from rm97.dbo.ReportSchedulerReports where Lower(Report) like '" +
                    MiscFunctions.HandleApostrophe(TrimmedReportFileName).ToLower() + "'", ref ErrMsg);
                    DateRangeString = dtSchedulerReports.Rows[0].Field<string>("DateRangeString");
                    DateRangeValue = dtSchedulerReports.Rows[0].Field<string>("DateRangeValue");
                    CrystalParameter1Name = dtSchedulerReports.Rows[0].Field<string>("CrystalParameter1Name");
                    CrystalParameter1ValueType = dtSchedulerReports.Rows[0].Field<string>("CrystalParameter1ValueType");
                    CrystalParameter1Value = dtSchedulerReports.Rows[0].Field<string>("CrystalParameter1Value");
                    CrystalParameter2Name = dtSchedulerReports.Rows[0].Field<string>("CrystalParameter2Name");
                    CrystalParameter2ValueType = dtSchedulerReports.Rows[0].Field<string>("CrystalParameter2ValueType");
                    CrystalParameter2Value = dtSchedulerReports.Rows[0].Field<string>("CrystalParameter2Value");
                    ExportType = dtSchedulerReports.Rows[0].Field<string>("ExportType");

                    //Load Report and set Parameters...
                    ReportDocument rptDoc = new ReportDocument();
                    rptDoc.Load(FullReportFileName);

                    //Set DateRange Parameters if they exist....
                    foreach (ParameterField fld in rptDoc.ParameterFields)
                    {
                        //DateRange Parameters found in Report so set them...
                        if (fld.ParameterValueType == ParameterValueKind.DateParameter |
                            fld.ParameterValueType == ParameterValueKind.DateTimeParameter)
                        {
                            ParameterRangeValue RangeVal = new ParameterRangeValue();
                            CalcStartAndEndDates(DateRangeString, DateRangeValue, ref StartDate, ref EndDate);
                            RangeVal.StartValue = StartDate;
                            RangeVal.EndValue = EndDate;
                            rptDoc.SetParameterValue(fld.Name, RangeVal);
                            fld.HasCurrentValue = true;
                        }
                        //====================================================================================
                        //Set Non-DateRange Parameters (Up to maximum of 2)...
                        //this is used to replace recordselection if parameter is (SELECT ALL)
                        string StringToMatch = "= {?" + fld.Name + "}";
                        string[] pStringArray = null;
                        string ParamName = "";
                        string ParamValue = "";

                        if (fld.Name == CrystalParameter1Name)
                        {
                            ParamValue = CrystalParameter1Value;
                            ParamName = CrystalParameter1Name;
                            if (ParamValue == "SELECT (ALL)")
                            {
                                rptDoc.RecordSelectionFormula = rptDoc.RecordSelectionFormula.Replace(StringToMatch, "<>''");
                            }
                            else
                            {
                                if (ParamValue.Contains(";"))
                                {
                                    pStringArray = ParamValue.Split(';');
                                    rptDoc.SetParameterValue(ParamName, pStringArray);
                                }
                                else
                                {
                                    rptDoc.SetParameterValue(ParamName, ParamValue);
                                }
                            }

                        }

                        if (fld.Name == CrystalParameter2Name)
                        {
                            ParamValue = CrystalParameter2Value;
                            ParamName = CrystalParameter2Name;
                            if (ParamValue == "SELECT (ALL)")
                            {
                                rptDoc.RecordSelectionFormula = rptDoc.RecordSelectionFormula.Replace(StringToMatch, "<>''");
                            }
                            else
                            {
                                if (ParamValue.Contains(";"))
                                {
                                    pStringArray = ParamValue.Split(';');
                                    rptDoc.SetParameterValue(ParamName, pStringArray);
                                }
                                else
                                {
                                    rptDoc.SetParameterValue(ParamName, ParamValue);
                                }
                            }
                        }

                    }
                    if (!rptDoc.HasRecords)
                    {
                        ErrMsg = "Error: No data for criteria selected.";
                        return false;
                    }
                    //Otherwise, report is ready to export...
                    //Export Report....
                    string attachfilename = Path.GetFileName(rptDoc.FileName).ToLower();
                    string ExportFile = "";
                    bool ExcelExport = false;
                    string ext = "";
                    if (ExportType == "Excel")
                    {
                        ExcelExport = true;
                        ext = ".xlsx";
                    }
                    else
                    {
                        ext = ".pdf";
                    }
                    ExportFile = MiscFunctions.GetUniqueTempPath() + MiscFunctions.ProperCase(Path.GetFileName(attachfilename)).Replace(".Rpt", ext);
                    AttachedFiles.Add(ExportFile);
                    if (ExcelExport)
                    {
                        ExportOptions exportOpts = new ExportOptions();
                        exportOpts = rptDoc.ExportOptions;
                        exportOpts.ExportFormatType = ExportFormatType.ExcelWorkbook;
                        ExcelDataOnlyFormatOptions excelFormatOpts = new ExcelDataOnlyFormatOptions();
                        DiskFileDestinationOptions diskOpts = new DiskFileDestinationOptions();
                        ExcelFormatOptions exportFormatOptions = ExportOptions.CreateExcelFormatOptions();
                        exportOpts.ExportFormatOptions = exportFormatOptions;
                        exportOpts.ExportDestinationType = ExportDestinationType.DiskFile;
                        excelFormatOpts.ExportObjectFormatting = true;
                        excelFormatOpts.SimplifyPageHeaders = true;
                        excelFormatOpts.MaintainRelativeObjectPosition = true;
                        diskOpts.DiskFileName = ExportFile;
                        exportOpts.DestinationOptions = diskOpts;
                        ClearReportHeaderAndFooter(rptDoc);
                        rptDoc.Export();
                        MiscFunctions.AutoFitAndFreezeTopRow(ExportFile);
                    }
                    else
                    {
                        rptDoc.ExportToDisk(ExportFormatType.PortableDocFormat, ExportFile);
                    }
                    MiscFunctions.LogMsg("Report: " + ReportFileName + " -Processed.", LogFile);
                    rptDoc.Close();
                }

                string EmailToStr = "";
                string EmailCCStr = "";
                string EmailSubjectStr = "";

                if (RunString == "RunSchedulerProfileTest")
                {
                    //Test run -email will be sent to tester only (args[3])
                    EmailToStr = args[3];
                    EmailCCStr = "";
                    EmailSubjectStr = "(Testing Profile) " + dtProfile.Rows[0].Field<string>("EmailSubject");

                }
                else//normal run...
                {
                    EmailToStr = dtProfile.Rows[0].Field<string>("EmailTo");
                    EmailCCStr = dtProfile.Rows[0].Field<string>("EmailCC");
                    EmailSubjectStr = dtProfile.Rows[0].Field<string>("EmailSubject");
                }

                //==Send Email===================================================================================
                SendEmail(
                "rcmsreports@prezero.us",
                EmailToStr,
                EmailCCStr,
                "",
                "rcmsreports@prezero.us",
                MiscFunctions.GetBracketFields(EmailSubjectStr, StartDate, EndDate),
                MiscFunctions.GetBracketFields(dtProfile.Rows[0].Field<string>("EmailBody"), StartDate, EndDate),
                ConcatStringArray(AttachedFiles, ";"), true, "smtp.office365.com", 587, "rcmsreports@prezero.us", EmailPswd,
                ref ErrMsg);
                //================================================================================================

                MiscFunctions.DeleteTempAttachments(AttachedFiles);

                if (ErrMsg != "")
                {
                    ErrMsg = "Error sending email: " + ErrMsg;
                    return false;
                }

                if (RunString == "RunSchedulerProfileTest")
                {
                    TsFunctions.DeleteTask(args[4]);//taskname (args[4])
                }
                else
                {
                    MiscFunctions.RunGenericSQLCmd("update rm97.dbo.ReportSchedulerProfiles set LastRunDate=GETDATE() where ProfileID=" + args[2], ref ErrMsg);
                }

                return true;
            }
            catch (Exception ex)
            {
                MiscFunctions.LogMsg("Fatal Error: " + ex.Message, LogFile);
                MiscFunctions.DeleteTempAttachments(AttachedFiles);
                return false;
            }
        }
        public static void SetConnStringAndFolders()
        {
            if (Environment.MachineName == "RCMS01")
            {
                RM97_ConnectionStr = Properties.Settings.Default.RCMS_Prod;
                LogFile = @"\\RCMS01\C$\DOCS\RCMS\Logs\RCMS Log.txt";
                RCMSFolder = @"\\RCMS01\C$\DOCS\RCMS\";
                ProdOrDevPath = @"\\RCMS01\";
            }
            else if (Environment.MachineName == "RCMS-DEV")
            {
                RM97_ConnectionStr = Properties.Settings.Default.RCMS_Dev;
                LogFile = @"\\RCMS-DEV\C$\DOCS\RCMS\Logs\RCMS Log.txt";
                RCMSFolder = @"\\RCMS-DEV\C$\DOCS\RCMS\";
                ProdOrDevPath = @"\\RCMS-DEV\";
            }
            else
            {
                RM97_ConnectionStr = Properties.Settings.Default.RCMS_Local;
                LogFile = @"\\" + Environment.MachineName + @"\C$\DOCS\RCMS\Logs\RCMS Log.txt";
                RCMSFolder = @"\\" + Environment.MachineName + @"\C$\DOCS\RCMS\";
                ProdOrDevPath = @"\\" + Environment.MachineName + @"\";
            }

            ReportsFolder = LogFile.Replace(@"RCMS\Logs\RCMS Log.txt", @"Reports\");
        }
        public static string Quote(string text)
        {
            return (char)34 + text + (char)34;
        }
        public static bool IsNumeric(string value)
        {
            return int.TryParse(value, out int i);
        }
        public static string ConcatStringArray(ArrayList str_array, string deliminator)
        {
            string Concat = "";
            foreach (string str in str_array)
            {
                Concat = Concat + str + deliminator;
            }
            Concat = Concat.TrimEnd(deliminator[0]);
            return Concat;
        }
        static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            var resourceName = Assembly.GetExecutingAssembly().GetName().Name + ".Lib." + new AssemblyName(args.Name).Name + ".dll";
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    var assemblyData = new Byte[stream.Length];
                    stream.Read(assemblyData, 0, assemblyData.Length);
                    return Assembly.Load(assemblyData);
                }
                else
                {
                    return null;
                }
            }
        }
        public static void CalcStartAndEndDates(
            string DateRangeString,
            string DateRangeValue,
            ref DateTime StartDate,
            ref DateTime EndDate)
        {
            switch (DateRangeString)
            {
                case "YTD":
                    StartDate = Convert.ToDateTime(@"1/1/" + DateToday.Year.ToString() + " 00:00:00");
                    EndDate = DateTime.Now;
                    break;
                case "MTD":
                    StartDate = Convert.ToDateTime(DateToday.Month.ToString() + @"/1/" + DateToday.Year.ToString() + " 00:00:00");
                    EndDate = DateTime.Now;
                    break;
                case "WTD":
                    StartDate = Convert.ToDateTime(DateToday.GetPreviousWeekDay(DayOfWeek.Sunday, -1));
                    EndDate = DateTime.Now;
                    break;
                case "LFM":
                    StartDate = Convert.ToDateTime(DateToday.PreviousMonthFirstDay().ToString("MM/dd/yyyy") + " 00:00:00");
                    EndDate = Convert.ToDateTime(DateToday.PreviousMonthLastDay().ToString("MM/dd/yyyy") + " 23:59:59");
                    break;
                case "CustomDates":
                    string[] DateRange = DateRangeValue.Split(':');
                    StartDate = Convert.ToDateTime(DateRange[0]);
                    EndDate = Convert.ToDateTime(DateRange[1]);
                    break;
                case "CustomStart":
                    StartDate = Convert.ToDateTime(DateRangeValue);
                    EndDate = DateTime.Now;
                    break;
                case "DaysPlusMinus":
                    if (Convert.ToInt32(DateRangeValue) < 0)//backward number of days
                    {
                        StartDate = DateTime.Now.AddDays(Convert.ToInt32(DateRangeValue));
                        EndDate = Convert.ToDateTime(DateToday.ToString("MM/dd/yyyy 23:59:59"));
                    }
                    else//forward number of days
                    {
                        StartDate = DateTime.Now;
                        EndDate = DateTime.Now.AddDays(Convert.ToInt32(DateRangeValue));
                    }
                    break;
                case "WeeksPlusMinus":
                    StartDate = Helpers.GetPreviousWeekDay(DateTime.Now, DayOfWeek.Sunday, Convert.ToInt32(DateRangeValue));
                    EndDate = DateTime.Now;
                    break;
                case "MonthsPlusMinus":
                    if (Convert.ToInt32(DateRangeValue) < 0)//backward number of months
                    {
                        StartDate = DateTime.Now.AddMonths(Convert.ToInt32(DateRangeValue));
                        EndDate = DateTime.Now;
                    }
                    else//forward number of months
                    {
                        StartDate = DateTime.Now;
                        EndDate = DateTime.Now.AddMonths(Convert.ToInt32(DateRangeValue));
                    }
                    break;
            }
        }
        public static void ClearReportHeaderAndFooter(ReportDocument rpt)
        {
            foreach (Section section in rpt.ReportDefinition.Sections)
            {
                if (section.Kind == AreaSectionKind.ReportHeader || section.Kind == AreaSectionKind.ReportFooter || section.Kind == AreaSectionKind.PageFooter)
                {
                    section.SectionFormat.EnableSuppress = true;
                    section.SectionFormat.BackgroundColor = Color.White;
                    foreach (var repO in section.ReportObjects)
                    {
                        if (repO is ReportObject)
                        {
                            var reportObject = repO as ReportObject;
                            reportObject.ObjectFormat.EnableSuppress = true;

                            reportObject.Border.BorderColor = Color.White;
                        }
                    }
                }
            }
        }
        public static void SendEmail(string strFrom, string strTo, string strCC, string strBC,
        string strReplyTo, string strSubject, string strBody, string strAttachment, bool IsBodyHTML,
        string strSmtpClient, int intSmtpPort, string strLoginName, string strPassword, ref string ErrMsg)
        {
            MailMessage mm = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            try
            {

                if (strTo != null & strTo != "")
                {
                    strTo = strTo.TrimEnd(';');
                    strTo = strTo.TrimEnd('\r', '\n');
                }
                if (strCC != null & strCC != "")
                {
                    strCC = strCC.TrimEnd(';');
                    strCC = strCC.TrimEnd('\r', '\n');
                }
                if (strBC != null & strBC != "")
                {
                    strBC = strBC.TrimEnd(';');
                    strBC = strBC.TrimEnd('\r', '\n');
                }

                if (strTo == null) { strTo = ""; }
                if (strCC == null) { strCC = ""; }
                if (strBC == null) { strBC = ""; }
                if (strReplyTo == null) { strReplyTo = ""; }
                if (strAttachment == null) { strAttachment = ""; }

                string[] EmailTo = strTo.Split(';');
                string[] EmailCC = strCC.Split(';');
                string[] EmailBC = strBC.Split(';');
                string[] Attachment = strAttachment.Split(';');


                mm.From = new MailAddress(strFrom);
                mm.Subject = strSubject;
                mm.Body = strBody;
                mm.IsBodyHtml = IsBodyHTML;



                if (strTo.Trim() != "")
                {
                    foreach (string Str in EmailTo) { mm.To.Add(new MailAddress(Str)); }
                }
                if (strCC.Trim() != "")
                {
                    foreach (string Str in EmailCC) { mm.CC.Add(new MailAddress(Str)); }
                }
                if (strBC.Trim() != "")
                {
                    foreach (string Str in EmailBC) { mm.Bcc.Add(new MailAddress(Str)); }
                }
                if (strReplyTo.Trim() != "")
                {
                    mm.ReplyToList.Add(new MailAddress(strReplyTo));
                }

                foreach (string Str in Attachment)
                {
                    if (Str == "")
                    {
                        break;
                    }
                    Attachment attachFile = new Attachment(Str);
                    mm.Attachments.Add(attachFile);
                }

                smtp.UseDefaultCredentials = false;
                smtp.Host = strSmtpClient;
                smtp.EnableSsl = true;
                smtp.Port = intSmtpPort;
                smtp.Timeout = 600000;
                System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential();
                NetworkCred.UserName = strLoginName;
                NetworkCred.Password = strPassword;
                smtp.Credentials = NetworkCred;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;


                smtp.Send(mm);

                mm.Dispose();
                smtp = null;
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                mm.Dispose();
                smtp = null;
            }
        }

    }
    public static class Helpers
    {
        public static DateTime PreviousMonthFirstDay(this DateTime currentDate)
        {
            DateTime d = currentDate.PreviousMonthLastDay();

            return new DateTime(d.Year, d.Month, 1);
        }

        public static DateTime PreviousMonthLastDay(this DateTime currentDate)
        {
            return new DateTime(currentDate.Year, currentDate.Month, 1).AddDays(-1);
        }

        public static DateTime GetPreviousWeekDay(this DateTime currentDate, DayOfWeek dow, int AddDays)
        {
            int CalcAddDays = 0;
            int currentDay = (int)currentDate.DayOfWeek, gotoDay = (int)dow;
            if (AddDays < 0)
            {
                CalcAddDays = (AddDays * -1) * 7;
                return currentDate.AddDays(CalcAddDays * -1).AddDays(gotoDay - currentDay);
            }
            else
            {
                CalcAddDays = AddDays * 7;
                return currentDate.AddDays(CalcAddDays).AddDays(gotoDay - currentDay);
            }
        }
    }
    public static class DayFinder
    {
        //For example to find the day for 2nd Friday, February, 2016
        //=>call FindDay(2016, 2, DayOfWeek.Friday, 2)
        public static int FindDay(int year, int month, DayOfWeek Day, int occurance)
        {

        start:
            if (occurance == 0 || occurance > 5)
                throw new Exception("Occurance is invalid");

            DateTime firstDayOfMonth = new DateTime(year, month, 1);
            //Substract first day of the month with the required day of the week 
            var daysneeded = (int)Day - (int)firstDayOfMonth.DayOfWeek;
            //if it is less than zero we need to get the next week day (add 7 days)
            if (daysneeded < 0) daysneeded = daysneeded + 7;
            //DayOfWeek is zero index based; multiply by the Occurance to get the day
            var resultedDay = (daysneeded + 1) + (7 * (occurance - 1));

            if (resultedDay > (firstDayOfMonth.AddMonths(1) - firstDayOfMonth).Days)
            {
                //throw new Exception(String.Format("No {0} occurance of {1} in the required month", occurance, Day.ToString()));
                occurance = occurance - 1; goto start;
            }
            return resultedDay;
        }
    }
}

