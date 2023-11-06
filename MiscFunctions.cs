using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Reflection;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace RCMSTaskScheduler.GlobalCode
{
    public class MiscFunctions
    {
        public static string RM97_ConnectionStr = Program.RM97_ConnectionStr;

        public static bool RunStoredProc(string proc, SqlParameter[] parameters, ref string ErrMsg)
        {

            try
            {
                System.Windows.Forms.Application.DoEvents();
                SqlConnection myConn = new SqlConnection(RM97_ConnectionStr);
                myConn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = proc;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Connection = myConn;
                    cmd.Parameters.AddRange(parameters);
                    cmd.ExecuteNonQuery();
                }
                myConn.Close();
                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                return false;
            }
        }

        public static void AutoFitAndFreezeTopRow(string xlsxFile)
        {
            Application xlApp = new Application();
            xlApp.DisplayAlerts = false;
            Workbook xlWorkBook = xlApp.Workbooks.Open(xlsxFile, 0,false, 5, "", "", 
                false, XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Columns.AutoFit();
            xlWorkSheet.Rows.AutoFit();
            xlWorkSheet.Activate();
            xlWorkSheet.Application.ActiveWindow.SplitRow = 1;
            xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
            xlWorkBook.Save();
            xlWorkBook.Close();
            xlApp.Quit();
        }

        public static int GetGalCount()
        {
            string ErrMsg = "";
            DataTable dt = RunGenericSelProc("select count(*) as GalCount from rm97.dbo.GAL", ref ErrMsg);
            return dt.Rows[0].Field<int>("GalCount");
        }

        public static bool AddressExist(string address)
        {
            string ErrMsg = "";
            DataTable dt = RunGenericSelProc("select EmailAddress from RM97.dbo.GAL where EmailAddress='"+ address + "'", ref ErrMsg);
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static string GetPswd(string Passwordtype)
        {
            string ErrMsg = "";
            System.Data.DataTable dt = RunGenericSelProc("Select password from rm97.dbo.CREDS where Type='" + Passwordtype + "'", ref ErrMsg);
            return dt.Rows[0].Field<string>("password");
        }
        public static string Decrypt(string cipherText, string E_Key)
        {
            string EncryptionKey = E_Key;
            cipherText = cipherText.Replace(" ", "+");
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
            0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76
        });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }
        public static string HandleApostrophe(string inpString)
        {
            return inpString.Replace("'", "''");
        }
        public static string GetEmailPassWord()
        {
            string ErrMsg = "";
            DataTable dtKey = RunGenericSelProc("select CurrentKey from rm97.dbo.E_KEY", ref ErrMsg);
            string Ekey = dtKey.Rows[0].Field<string>("CurrentKey");
            string EmailPswd = Decrypt(GetPswd("EmailPswd"), Ekey);
            return EmailPswd;
        }
        public static string GetTSPassWord()
        {
            string ErrMsg = "";
            DataTable dtKey = RunGenericSelProc("select CurrentKey from rm97.dbo.E_KEY", ref ErrMsg);
            string Ekey = dtKey.Rows[0].Field<string>("CurrentKey");
            string EmailPswd = Decrypt(GetPswd("TS_Pswd"), Ekey);
            return EmailPswd;
        }


        public static string GetUniqueTempPath()
        {
            string uniquePath;
            do
            {
                Guid guid = Guid.NewGuid();
                string uniqueSubFolderName = guid.ToString();
                uniquePath = Path.GetTempPath() + uniqueSubFolderName;
            } while (Directory.Exists(uniquePath));
            Directory.CreateDirectory(uniquePath);
            return uniquePath + @"\";
        }

        public static bool RunGenericSQLCmd(string SqlCmd, ref string ErrMsg)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(RM97_ConnectionStr);
                SqlCommand cmd = new SqlCommand(SqlCmd, myConn);
                myConn.Open();
                cmd.ExecuteScalar();
                myConn.Close();
                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                return false;
            }
        }
        public static DataTable RunGenericSelProc(string proc, ref string ErrMsg)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(RM97_ConnectionStr);
                DataSet data = new DataSet();
                myConn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = proc;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = myConn;
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        da.Fill(data);
                    }
                }
                DataTable dt = data.Tables[0];
                myConn.Close();
                return dt;
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;
                return null;
            }
        }
        public static DataTable GetProfile(string ProfileID)
        {
            string ErrMsg = "";
            DataTable dt = RunGenericSelProc("select * from rm97.dbo.ReportSchedulerProfiles where ProfileID=" +
            ProfileID, ref ErrMsg);

            if (ErrMsg != "")
            {
                LogMsg("Error: " + ErrMsg, Program.LogFile);
            }
            return dt;
        }

        public static void LogMsg(String logmsg, string logfile)
        {
            if (logfile != "")
            {
                StreamWriter w = File.AppendText(logfile);
                w.WriteLine(logmsg);
                w.Flush();
                w.Close();
            }
        }
        public static void LogInitialize(string logfile, string logmsg, ref string ErrMsg)
        {
            try
            {
                string LogFilePath = Path.GetDirectoryName(logfile);
                if (!Directory.Exists(LogFilePath))
                {
                    Directory.CreateDirectory(LogFilePath);
                }
                StreamWriter w = File.AppendText(logfile);
                w.WriteLine("_______________________________________________");
                w.WriteLine("Log Entry:  " + DateTime.Now.ToString());
                w.WriteLine("{0}", logmsg);
                w.Flush();
                w.Close();
            }
            catch (Exception ex)
            {
                ErrMsg = ex.Message;

            }
        }
        public static string GetReportPathFromFileName(string fileName)
        {
            //returns the report path where report filename is found... 
            //(this way we can move reports to other folders without breaking the system)
            string pathFound = "";
            try
            {
                foreach (string path in Directory.EnumerateFiles(Program.ReportsFolder, fileName, SearchOption.AllDirectories))
                {
                    pathFound = Path.GetDirectoryName(path) + @"\";
                    break;
                }
                return pathFound;
            }
            catch (Exception)
            {
                return "";
            }
        }

        public static string ProperCase(string str)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string PropercaseText = str.ToLower();
            PropercaseText = textInfo.ToTitleCase(PropercaseText);
            return PropercaseText;

        }
        public static string GetBracketFields(string MsgBodyOrSubject, DateTime StartDate, DateTime EndDate)
        {
            MatchCollection matches = GetTextInBrackets(MsgBodyOrSubject);
            foreach (Match m in matches)
            {
                switch (m.Groups[1].Value.ToString().ToLower())
                {
                    case "startdate":
                        MsgBodyOrSubject = MsgBodyOrSubject.Replace("[startdate]", StartDate.ToString("MM-dd-yy"));
                        break;
                    case "enddate":
                        MsgBodyOrSubject = MsgBodyOrSubject.Replace("[enddate]", EndDate.ToString("MM-dd-yy"));
                        break;
                    case "day":
                        MsgBodyOrSubject = MsgBodyOrSubject.Replace("[day]", StartDate.ToString("dddd"));
                        break;
                    case "month":
                        MsgBodyOrSubject = MsgBodyOrSubject.Replace("[month]", StartDate.ToString("MMMM"));
                        break;
                }
            }
            return MsgBodyOrSubject;
        }
        public static MatchCollection GetTextInBrackets(string inputString)
        {
            return Regex.Matches(inputString, @"\[(.*?)\]");
        }

        public static void DeleteTempAttachments(ArrayList AttachedFiles)
        {
            try
            {
                foreach (string str in AttachedFiles)
                {
                    string folderToDelete = Directory.GetParent(str).ToString();
                    Directory.Delete(folderToDelete, true);
                }
            }
            catch (Exception)
            {

            }
        }
    }
}
