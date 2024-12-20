using System;
using System.Data;
using System.IO;
using System.IO.Pipes;
using System.Text.RegularExpressions;
using System.Net;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using DataTable = System.Data.DataTable;
using System.Text.Json;
//using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using static System.Net.Mime.MediaTypeNames;
using System.Security.Policy;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualBasic.Logging;
using System.Runtime.ConstrainedExecution;

namespace API
{
    [DisplayName("Jira Defect Data")]
    public class DefectData
    {
        [DisplayName("Creation Date")]
        public string Creation_Date { get; set; } = "";
        [DisplayName("Defect Key")]
        public string Key { get; set; } = "";
        [DisplayName("Sprint")]
        public string Sprint { get; set; } = "";
        [DisplayName("Sprint Start Date")]
        public string Sprint_Start_Date { get; set; } = "";
        [DisplayName("Sprint End Date")]
        public string Sprint_End_Date { get; set; } = "";
        [DisplayName("Defect Status")]
        public string Status { get; set; } = "";
        [DisplayName("Defect Summary")]
        public string Summary { get; set; } = "";
        [DisplayName("Defect Type")]
        public string Type { get; set; } = "";
        [DisplayName("Issue Type")]
        public string Issue_Type { get; set; } = "";
        [DisplayName("Assignee")]
        public string Assignee { get; set; } = "";
        [DisplayName("Defect Priority")]
        public string Priority { get; set; } = "";
        [DisplayName("Defect Severity")]
        public string Severity { get; set; } = "";
        [DisplayName("Defect Resolution")]
        public string Resolution { get; set; } = "";
        [DisplayName("Defect Purpose")]
        public string Purpose { get; set; } = "";
    }
    public class TestExecutionReport
    {
       // public string 
    }
    
    public class TEMPLINKVAL
    {
        public string url { get; set; } = "";
        public string value { get; set; } = "";
    }
    public static class JIRAALLDefects
    { 
        public static List<TEMPLINKVAL> TLIST = new();
        public static void Main()
        {
            
            try
            {
                Console.WriteLine("Time Start: " + DateTime.Now.ToString("yyyy-MM-dd Hh24:mm:ss"));
                //List<string> SERVERLIST = GETNEOLOADSERVERS();
                //GENROCKETDATA();
                JIRADEFECTdata();
                //NEOLOADDATA(SERVERLIST);
                GETZEPHYRSCALEDATATESTCASE();
                GETZEPHYRSCALEDATATESTEXECUTION();
                Console.WriteLine("Time End: " + DateTime.Now.ToString("yyyy-MM-dd Hh24:mm:ss"));
            }
            catch
            {
            }
        }
        public class TestCaseData
        {
            [DisplayName("Test Case")]
            public string TestCase { get; set; } = "";
            [DisplayName("Project")]
            public string Project { get; set; } = "";
            [DisplayName("Created On")]
            public string CreatedOn { get; set; } = "";
            public string Purpose { get; set; } = "";
        }
        public class TextExecutionData
        {
            [DisplayName("Test Execution Key")]
            public string TestExecutionKey { get; set; } = "";
            [DisplayName("Project")]
            public string Project { get; set; } = "";
            [DisplayName("Test Case")]
            public string TestCase { get; set; } = "";
            [DisplayName("Test Cycle")]
            public string TestCycle { get; set; } = "";
            [DisplayName("Test Environment")]
            public string Environment { get; set; } = "";
            [DisplayName("Assigned To")]
            public string AssignedTo { get; set; } = "";
            [DisplayName("Actual End Date")]
            public string ActualEndDate { get; set; } = "";
            [DisplayName("_time")]
            public string _time { get; set; } = "";
            [DisplayName("Test Status")]
            public string Status { get; set; } = "";
            [DisplayName("Purpose")]
            public string Purpose { get; set; } = "";
            [DisplayName("Execution Type")]
            public string ExecutionType { get; set; } = "";
        }
        public static void GETZEPHYRSCALEDATATESTEXECUTION()
        {
            try
            {
                int startindx = 0;
                int maxresults = 50;
                bool continueloop = true;
                Console.WriteLine("Which Project you need the defect details on: ");
                string Prj = Console.ReadLine();
                DataTable ZSCALEEXEC = new();
                ZSCALEEXEC.Columns.Add("Test Execution Key");
                ZSCALEEXEC.Columns.Add("Project");
                ZSCALEEXEC.Columns.Add("Test Case");
                ZSCALEEXEC.Columns.Add("Test Cycle");
                ZSCALEEXEC.Columns.Add("Environment");
                ZSCALEEXEC.Columns.Add("Assigned To");
                ZSCALEEXEC.Columns.Add("_time");
                ZSCALEEXEC.Columns.Add("Actual End Date");
                ZSCALEEXEC.Columns.Add("Execution Type");
                ZSCALEEXEC.Columns.Add("Status");
                List<TextExecutionData> TEArray = new List<TextExecutionData>();
                while (continueloop)
                {
                    string _URL = "https://api.zephyrscale.smartbear.com/v2/testexecutions?projectKey=" + Prj + "&maxResults=" + maxresults.ToString() + "&startAt=" + startindx.ToString();

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                    request.Method = "GET";
                    request.UserAgent = "AEQE Analysis";
                    request.Headers.Add("Authorization", GetAuth("ZEPHYRAUTH"));
                    request.Headers.Add("Accept", "application/json");
                    request.Headers.Add("Content-Type", "application/json");
                    request.KeepAlive = true;
                    var response = (HttpWebResponse)request.GetResponse();
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                    JObject jobj = JObject.Parse(responseString);
                    JArray jArray = (JArray)jobj["values"];
                    int jsize = jArray.Count;

                    for (int i = 0; i < jsize; i++)
                    {
                        TextExecutionData teitem = new TextExecutionData();
                        teitem.Purpose = Prj + " Test Execution Data";
                        DataRow DR = ZSCALEEXEC.NewRow();
                        DR["Test Execution Key"] = jobj["values"][i]["key"].ToString();
                        teitem.TestExecutionKey = jobj["values"][i]["key"].ToString();
                        try
                        {
                            DR["Project"] = GetProjectData(jobj["values"][i]["project"]["self"].ToString(), jobj["values"][i]["project"]["id"].ToString());
                            teitem.Project = GetProjectData(jobj["values"][i]["project"]["self"].ToString(), jobj["values"][i]["project"]["id"].ToString());
                        }
                        catch
                        {
                        }
                        DR["Test Case"] = GetProjectData(jobj["values"][i]["testCase"]["self"].ToString(), jobj["values"][i]["testCase"]["id"].ToString(), "Test Case");
                        teitem.TestCase = GetProjectData(jobj["values"][i]["testCase"]["self"].ToString(), jobj["values"][i]["testCase"]["id"].ToString(), "Test Case");
                        DR["Status"] = GetProjectData(jobj["values"][i]["testExecutionStatus"]["self"].ToString(), jobj["values"][i]["testExecutionStatus"]["id"].ToString(), "Status");
                        teitem.Status = GetProjectData(jobj["values"][i]["testExecutionStatus"]["self"].ToString(), jobj["values"][i]["testExecutionStatus"]["id"].ToString(), "Status");
                        try
                        {
                            DR["Environment"] = GetProjectData(jobj["values"][i]["environment"]["self"].ToString(), jobj["values"][i]["environment"]["id"].ToString(), "Environment");
                            if(string.IsNullOrEmpty(Convert.ToString(DR["Environment"])))
                            {
                                DR["Environment"] = "Not Mentioned";
                            }
                            teitem.Environment = Convert.ToString(DR["Environment"]);
                        }
                        catch
                        {
                            DR["Environment"] = "Not Mentioned";
                            teitem.Environment = "Not Mentioned";
                        }
                        DR["Actual End Date"] = jobj["values"][i]["actualEndDate"].ToString().Split(' ')[0];
                        teitem.ActualEndDate = jobj["values"][i]["actualEndDate"].ToString().Split(' ')[0];

                        DR["_time"] = jobj["values"][i]["actualEndDate"].ToString().Split(' ')[0];
                        teitem._time = jobj["values"][i]["actualEndDate"].ToString().Split(' ')[0];

                        DR["Execution Type"] = jobj["values"][i]["automated"].ToString();
                        teitem.ExecutionType= jobj["values"][i]["automated"].ToString();
                        try
                        {
                            if (!string.IsNullOrEmpty(jobj["values"][i]["assignedToId"].ToString()))
                            {
                               // Console.WriteLine("https://fish-net.atlassian.net/rest/api/3/user?accountId=" + jobj["values"][i]["assignedToId"].ToString());
                                DR["Assigned To"] = GetProjectData("https://fish-net.atlassian.net/rest/api/3/user?accountId=" + jobj["values"][i]["assignedToId"].ToString(), jobj["values"][i]["assignedToId"].ToString(), "Assigned To");
                                
                            }
                            if (string.IsNullOrEmpty(Convert.ToString(DR["Assigned To"])))
                            {
                                DR["Assigned To"] = "Unassigned";
                            }
                            teitem.AssignedTo = Convert.ToString(DR["Assigned To"]);
                        }
                        catch
                        {
                            DR["Assigned To"] = "Unassigned";
                            teitem.AssignedTo = "Unassigned";
                        }
                        try
                        {
                            DR["Test Cycle"] = GetProjectData(jobj["values"][i]["testCycle"]["self"].ToString(), jobj["values"][i]["environment"]["id"].ToString(), "Test Cycle");
                            teitem.TestCycle = GetProjectData(jobj["values"][i]["testCycle"]["self"].ToString(), jobj["values"][i]["environment"]["id"].ToString(), "Test Cycle");
                        }
                        catch
                        {
                            DR["Test Cycle"] = "Not Specified";
                            teitem.TestCycle = "Not Specified";
                        }
                        ZSCALEEXEC.Rows.Add(DR);
                        TEArray.Add(teitem);
                    }
                    if (jsize < 50)
                    {
                        continueloop = false;
                    }
                    startindx = startindx + 50;


                    request.Abort();
                    response.Close();

                }

                // MessageBox.Show(ZSCALEDATA.Rows.Count.ToString());

                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Sheets[1];
                worksheet.Name = "Test Case Data";

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                var columns = ZSCALEEXEC.Columns.Count;
                var rows = ZSCALEEXEC.Rows.Count;

                Excel.Range range = worksheet.Range["A1", String.Format("{0}{1}", GetExcelColumnName(columns), rows + 1)];

                object[,] data = new object[rows + 1, columns];

                for (int rowNumber = 0; rowNumber < rows; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                        {
                            data[rowNumber, columnNumber] = (ZSCALEEXEC.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                    {
                        data[rowNumber + 1, columnNumber] = Convert.ToString(ZSCALEEXEC.Rows[rowNumber][columnNumber]);
                    }
                }
                range.Value = data;

                if (File.Exists(@"C:\km\TExecDetails_" + Prj + ".xlsx"))
                {
                    File.Delete(@"C:\km\TExecDetails_" + Prj + ".xlsx");
                }
                workbook.SaveAs(@"C:\km\TExecDetails_" + Prj + ".xlsx");
                workbook.Close();
                Marshal.ReleaseComObject(application);

                var prettyJson = System.Text.Json.JsonSerializer.Serialize(TEArray, new JsonSerializerOptions { WriteIndented = true });
               
                File.WriteAllText(@"c:\km\" + Prj + "TestExecutionData.json", prettyJson);
                UploadToSPlunk(prettyJson);
            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message, "Error");
            }
        }
        public static void GETZEPHYRSCALEDATATESTCASE()
        {
            try
            {
                
                //getTestCaseData
                int startindx = 0;
                int maxresults = 50;
                Console.WriteLine("Which Project you need the Test Case details on: ");
                string Prj = Console.ReadLine();
                bool continueloop = true;
                DataTable ZSCALEDATA = new DataTable();
                ZSCALEDATA.Columns.Add("Test Case Key");
                ZSCALEDATA.Columns.Add("Test Case Name");
                ZSCALEDATA.Columns.Add("Created On");
                ZSCALEDATA.Columns.Add("Project");
                ZSCALEDATA.Columns.Add("Priority");
                ZSCALEDATA.Columns.Add("Status");
                ZSCALEDATA.Columns.Add("Purpose");
                List<TestCaseData> TCArray = new List<TestCaseData>();
                while (continueloop)
                {
                    string _URL = "https://api.zephyrscale.smartbear.com/v2/testcases?projectKey=" + Prj + "&maxResults=" + maxresults.ToString() + "&startAt=" + startindx.ToString();
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                    request.Method = "GET";
                    request.UserAgent = "AEQE Analysis";
                    request.Headers.Add("Authorization", GetAuth("ZEPHYRAUTH"));
                    request.Headers.Add("Accept", "application/json");
                    request.Headers.Add("Content-Type", "application/json");
                    request.KeepAlive = true;
                    var response = (HttpWebResponse)request.GetResponse();
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                    request.Abort();
                    response.Close();

                    JObject jobj = JObject.Parse(responseString);
                    JArray jArray = (JArray)jobj["values"];
                    int jsize = jArray.Count;

                    for (int i = 0; i < jsize; i++)
                    {
                        DataRow DR = ZSCALEDATA.NewRow();
                        TestCaseData TCD = new();
                        DR["Test Case Key"] = jobj["values"][i]["key"].ToString();
                        TCD.TestCase = jobj["values"][i]["key"].ToString();
                        DR["Purpose"]= Prj + " Test Case Data";
                        TCD.Purpose = Prj + " Test Case Data";
                        DR["Test Case Name"] = jobj["values"][i]["name"].ToString();
                        try
                        {
                            DR["Project"] = GetProjectData(jobj["values"][i]["project"]["self"].ToString(),jobj["values"][i]["project"]["id"].ToString());
                            TCD.Project = Convert.ToString(DR["Project"]);
                        }
                        catch
                        {
                        }

                        DR["Created On"] = jobj["values"][i]["createdOn"].ToString().Split(' ')[0]; 
                        TCD.CreatedOn = jobj["values"][i]["createdOn"].ToString().Split(' ')[0];
                        DR["Priority"] = GetProjectData(jobj["values"][i]["priority"]["self"].ToString(),jobj["values"][i]["priority"]["id"].ToString(), "Priority");
                        DR["Status"] = GetProjectData(jobj["values"][i]["status"]["self"].ToString(),jobj["values"][i]["status"]["id"].ToString(),"Status");
                        ZSCALEDATA.Rows.Add(DR);
                        TCArray.Add(TCD);
                    }
                    if (jsize <50)
                    {
                        continueloop = false;
                    }
                    startindx = startindx + 50;
                }
                // MessageBox.Show(ZSCALEDATA.Rows.Count.ToString());

                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Sheets[1];
                worksheet.Name = "Test Case Data";
                
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                var columns = ZSCALEDATA.Columns.Count;
                var rows = ZSCALEDATA.Rows.Count;

                Excel.Range range = worksheet.Range["A1", String.Format("{0}{1}", GetExcelColumnName(columns), rows + 1)];

                object[,] data = new object[rows + 1, columns];

                for (int rowNumber = 0; rowNumber < rows; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                        {
                            data[rowNumber, columnNumber] = (ZSCALEDATA.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                    {
                        data[rowNumber + 1, columnNumber] = Convert.ToString(ZSCALEDATA.Rows[rowNumber][columnNumber]);
                    }
                }
                range.Value = data;

                if (File.Exists(@"C:\km\TCDetails_" + "AEQE" + ".xlsx"))
                {
                    File.Delete(@"C:\km\TCDetails_" + "AEQE" + ".xlsx");
                }
                workbook.SaveAs(@"C:\km\TCDetails_" + "AEQE" + ".xlsx");
                workbook.Close();
                Marshal.ReleaseComObject(application);
                var prettyJson = System.Text.Json.JsonSerializer.Serialize(TCArray, new JsonSerializerOptions { WriteIndented = true });

                //File.WriteAllText(@"c:\km\" + Prj + "TestExecutionData.json", prettyJson);
                UploadToSPlunk(prettyJson);

            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message);
            }
        }
        public static string GetProjectData(string _URL,string prjID,string refstr="Project")
        {
            try
            {
                try
                {

                    if(TLIST.FindIndex(e => e.url.Equals(_URL))>=0)
                    {
                        return TLIST[TLIST.FindIndex(e => e.url.Equals(_URL))].value;
                    }

                }
                catch
                {

                }
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                request.Method = "GET";
                request.UserAgent = "AEQE Analysis";
                if (refstr != "Assigned To")
                {
                    request.Headers.Add("Authorization", GetAuth("ZEPHYRAUTH"));
                }
                else
                {
                    request.Headers.Add("Authorization", GetAuth("JIRAAUTH"));
                }
                request.Headers.Add("Accept", "application/json");
                request.Headers.Add("Content-Type", "application/json");
                request.KeepAlive = true;
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                request.Abort();
                response.Close();

                JObject jobj = JObject.Parse(responseString);
                TEMPLINKVAL tkl = new TEMPLINKVAL();
                tkl.url = _URL;
                if (refstr == "Project")
                {
                    tkl.value = jobj["key"].ToString();
                    TLIST.Add(tkl);
                    return jobj["key"].ToString();
                }
                else if (refstr =="Priority")
                {
                    tkl.value = jobj["name"].ToString();
                    TLIST.Add(tkl);
                    return jobj["name"].ToString();
                }
                else if (refstr == "Status")
                {
                    tkl.value = jobj["name"].ToString();
                    TLIST.Add(tkl);
                    return jobj["name"].ToString();
                }
                else if (refstr == "Environment")
                {
                    tkl.value = jobj["name"].ToString();
                    TLIST.Add(tkl);
                    return jobj["name"].ToString();
                }
                else if (refstr == "Test Case")
                {
                    tkl.value = jobj["name"].ToString();
                    TLIST.Add(tkl);
                    return jobj["name"].ToString();
                }
                else if (refstr == "Test Cycle")
                {
                    tkl.value = jobj["key"].ToString();
                    TLIST.Add(tkl);
                    return jobj["key"].ToString();
                }
                else if (refstr == "Assigned To")
                {
                    tkl.value = jobj["displayName"].ToString();
                    TLIST.Add(tkl);
                    return jobj["displayName"].ToString();
                }
                tkl.value = jobj["key"].ToString();
                TLIST.Add(tkl);
                return jobj["key"].ToString();
            }
            catch(Exception E)
            {
                return E.Message;
            }
        }
        public static List<string> GETNEOLOADSERVERS()
        {
            List<string> Prodservers = new();
            try
            {
                string _URL = "https://fish-net.atlassian.net//wiki/api/v2/pages/1132201180?body-format=ATLAS_DOC_FORMAT";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                request.Method = WebRequestMethods.Http.Get;

                request.UserAgent = "AEQE Analysis";
                //MessageBox.Show("1");
                request.Headers.Add("Authorization", GetAuth("JIRAAUTH"));
                request.Headers.Add("Accept", "application/json");
                request.Headers.Add("Content-Type", "application/json");
                request.KeepAlive = true;
                //MessageBox.Show("2");
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                //MessageBox.Show("3");
                string rx = @"([A-Za-z]{2})-([A-Za-z]{2})-([A-Za-z]{1})(\d{5}):7100";

                request.Abort();
                response.Close();
                MatchCollection m = Regex.Matches(responseString, rx, RegexOptions.IgnoreCase);
                foreach (Match match in m)
                {
                    Console.WriteLine("Server name : {0}, Index : {1}", match.Value, match.Index);
                    if(match.Value.ToString().ToUpper().StartsWith("CP"))
                    {
                        Prodservers.Add(match.Value.ToString().ToUpper());
                    }
                }
            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message, "Error");
            }
            return Prodservers;
        }
        public static void  NEOLOADDATA(List<string> SERVERLIST)
        {
            try
            {
                string _URL = "https://neoload-api.saas.neotys.com/v3/resources/zones";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                request.Method = WebRequestMethods.Http.Get;
                request.Accept = "appliation/json";
                request.ContentType = "application/json";
                request.UserAgent = "AEQE Analysis";
                //Request Headers
                request.Headers.Add("accountToken" , GetAuth("NEOLOADAUTH"));
                request.Headers.Add("Accept", "application/json");
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                request.Abort();
                response.Close();
                JArray jArray = JArray.Parse(responseString);
                DataTable DTAB = new DataTable();
                DTAB.Columns.Add("Resource Id");
                DTAB.Columns.Add("Resource Name");
                DTAB.Columns.Add("Resource Type");
                DTAB.Columns.Add("Load Controllers");
                DTAB.Columns.Add("Load Generators");

                int jsize = jArray.Count;

                for(int x=0;x<jsize;x++)
                {
                    DataRow DR          = DTAB.NewRow();
                    DR["Resource Id"]   = (jArray[0]["id"].ToString());
                    DR["Resource Name"] = (jArray[0]["name"].ToString());
                    DR["Resource Type"] = (jArray[0]["type"].ToString());                    
                    int controllerCount = jArray[x]["controllers"].Count();                    
                    DR["Load Controllers"] = "";                    
                    for (int y = 0;y<controllerCount;y++)
                    {
                        if (y > 0)
                        {
                            DR["Load Controllers"] = Convert.ToString(DR["Load Controllers"]) + "\r\n" + jArray[x]["controllers"][y]["name"].ToString();
                        }
                        else
                        {
                            DR["Load Controllers"] = jArray[x]["controllers"][y]["name"].ToString();
                        }
                    }
                    int loadgencount = jArray[x]["loadgenerators"].Count();
                    DR["Load Generators"] = "";

                    for (int y = 0; y < loadgencount; y++)
                    {
                        if (y > 0)
                        {
                            DR["Load Generators"] = Convert.ToString(DR["Load Generators"]) + "\r\n" + jArray[x]["loadgenerators"][y]["name"].ToString();
                        }
                        else
                        {
                            DR["Load Generators"] = jArray[x]["loadgenerators"][y]["name"].ToString();
                        }
                        if(SERVERLIST.Contains(jArray[x]["loadgenerators"][y]["name"].ToString().ToUpper()))
                        {
                            Console.WriteLine("Validated Server: " + jArray[x]["loadgenerators"][y]["name"].ToString().ToUpper());
                            SERVERLIST.Remove(jArray[x]["loadgenerators"][y]["name"].ToString().ToUpper());
                        }

                    }
                    DTAB.Rows.Add(DR);
                }
                if(SERVERLIST.Count >0)
                {
                    Console.WriteLine("Not All Production Servers are Validated");
                }
                else
                {
                    Console.WriteLine("All Production Servers found in confluence Page are Validated");
                }
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Sheets[1];
                worksheet.Name = "Neoload Resources";
                var columns = DTAB.Columns.Count;
                var rows = DTAB.Rows.Count;
                Excel.Range range = worksheet.Range["A1", String.Format("{0}{1}", GetExcelColumnName(columns), rows + 1)];
                object[,] data = new object[rows + 1, columns];
                for (int rowNumber = 0; rowNumber < rows; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                        {
                            data[rowNumber, columnNumber] = (DTAB.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                    {
                        data[rowNumber + 1, columnNumber] = Convert.ToString(DTAB.Rows[rowNumber][columnNumber]);
                    }
                }
                Console.WriteLine("....Range.....");
                range.Value = data;

                if (File.Exists(@"C:\km\NeoloadResources"  + ".xlsx"))
                {
                    File.Delete(@"C:\km\NeoloadResources" + ".xlsx");
                }
                workbook.SaveAs(@"C:\km\NeoloadResources" + ".xlsx");
                workbook.Close();
            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message);
            }
        }
        public static void GENROCKETDATA()
        {
            try
            {
                string _URL = "https://genrocketgmusvnextcloud.fi.com:9010/rest/scenario";
                string REQBODY = File.ReadAllText(@"C:\km\Body.json");
                Console.WriteLine("a1");
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls13;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
               // System.Net.ServicePointManager.ServerCertificateValidationCallback +=
               //delegate (object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
               //                        System.Security.Cryptography.X509Certificates.X509Chain chain,
               //                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
               //{
               //    try
               //    {
               //        Console.WriteLine("SSL certificate error: {0}", sslPolicyErrors);

               //        // bool certMatch = false; // Assume failure
               //        byte[] certHash = certificate.GetCertHash();
               //        //if (certHash.Length == apiCertHash.Length)
               //        //{
               //        //    certMatch = true; // Now assume success.
               //        //    for (int idx = 0; idx < certHash.Length; idx++)
               //        //    {
               //        //        if (certHash[idx] != apiCertHash[idx])
               //        //        {
               //        //            certMatch = false; // No match
               //        //            break;
               //        //        }
               //        //    }
               //        //}
               //        return true; // **** Always accept
               //    }
               //    catch
               //    {
               //        return true;
               //    }
               //};
                //Request Attributes
                Console.WriteLine("1c");
                request.Method = WebRequestMethods.Http.Post;
                request.Accept = "appliation/json";
                request.ContentType = "application/json";
                request.UserAgent = "CSharpCpde";
                request.Headers.Add("Accept", "application/json");
                byte[] data = System.Text.Encoding.ASCII.GetBytes(REQBODY);
                request.ContentType = "application/json";
                request.Accept = "application/json";
                request.ContentLength = data.Length;
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(data, 0, data.Length);
                Console.WriteLine("1d");
                requestStream.Close();
                var response = (HttpWebResponse)request.GetResponse();
                Console.WriteLine("1e");
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                request.Abort();
                Console.WriteLine(responseString.ToString());
            }
            catch(WebException E)
            {
                Console.WriteLine("Response Messge: " + E.Message.ToString());
                Console.WriteLine("Response Messge: " + E.InnerException.ToString());
                Console.WriteLine("Response: " + (E.Response as HttpWebResponse)?.StatusDescription);
                Console.WriteLine("Status Code: " + ((int)(E.Response as HttpWebResponse)?.StatusCode).ToString());
            }
        }
        public static void JIRADEFECTdata()
        {
            try
            {
                //URL
                Console.WriteLine("Which Project you need the defect details on: ");
                string Prj = Console.ReadLine();
                Console.WriteLine("Initializing Table Definitions.......");
                System.Data.DataTable DXT = new();
                //Add COlumns
                DXT.Columns.Add("Creation Date");
                DXT.Columns.Add("Key");
                DXT.Columns.Add("Sprint");
                DXT.Columns.Add("Sprint Start Date");
                DXT.Columns.Add("Sprint End Date");
                DXT.Columns.Add("Status");
                DXT.Columns.Add("Summary");
                DXT.Columns.Add("Type");
                DXT.Columns.Add("Issue Type");
                DXT.Columns.Add("Assignee");
                DXT.Columns.Add("Priority");
                DXT.Columns.Add("Severity");
                DXT.Columns.Add("Resolution");
                DXT.Columns.Add("Purpose");
                List<DefectData> DefectArray = new List<DefectData>();
                int stindx = 0;
                while (1 == 1)
                {
                    string _URL = "https://fish-net.atlassian.net/rest/api/2/search?jql=project=" + Prj + " And type=Bug order by created asc&fields=key,sprint,status,priority,issuetype,priority,assignee,summary,customfield_10344,created,resolution&startAt=" + stindx.ToString() + "&maxResults=100";
                    //Create Request
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                    //Request Attributes
                    request.Method = WebRequestMethods.Http.Get;
                    request.Accept = "appliation/json";
                    request.ContentType = "application/json";
                    request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0";
                    //Request Headers
                    request.Headers.Add("Authorization"
                                      , GetAuth("JIRAAUTH"));
                    request.Headers.Add("Accept", "application/json");
                    var response = (HttpWebResponse)request.GetResponse();
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                    request.Abort();
                    response.Close();
                    //json parsing
                    JObject jobj = JObject.Parse(responseString);
                    JArray jArray = (JArray)jobj["issues"];
                    Console.WriteLine("Defect Count : " + jArray.Count.ToString());
                    int jsize = jArray.Count;  
                    Console.WriteLine("Initializing Data Collections.......");
                    int j = 0, refint = jsize / 10;
                    for (int i = 0; i < jsize; i++)
                    {
                        DefectData DDa = new DefectData();
                        DataRow DXTRow = DXT.NewRow();
                        DXTRow["Key"] = jobj["issues"][i]["key"].ToString();
                        DDa.Key = jobj["issues"][i]["key"].ToString();
                        DXTRow["Status"] = jobj["issues"][i]["fields"]["status"]["name"].ToString();
                        DDa.Status = jobj["issues"][i]["fields"]["status"]["name"].ToString();
                        DXTRow["Summary"] = jobj["issues"][i]["fields"]["summary"].ToString();
                        DDa.Summary = jobj["issues"][i]["fields"]["summary"].ToString();
                        DXTRow["Type"] = "Bug";
                        DDa.Type = "Bug";
                        DXTRow["Purpose"] = Prj + " Defect Reporting";
                        DDa.Purpose = Prj + " Defect Reporting";
                        List<string> sdata = GetSprint(jobj["issues"][i]["key"].ToString());
                        DXTRow["Sprint"] = sdata[0];
                        DDa.Sprint = sdata[0];
                        DXTRow["Sprint Start Date"] = sdata[1];
                        DDa.Sprint_Start_Date = sdata[1];
                        DXTRow["Sprint End Date"] = sdata[2];
                        DDa.Sprint_End_Date = sdata[2];
                        DXTRow["Issue Type"] = jobj["issues"][i]["fields"]["issuetype"]["name"].ToString();
                        DDa.Issue_Type = jobj["issues"][i]["fields"]["issuetype"]["name"].ToString();
                        DXTRow["Assignee"] = Convert.ToString(jobj["issues"][i]["fields"]["Assignee"]);
                        DDa.Assignee = Convert.ToString(jobj["issues"][i]["fields"]["Assignee"]);
                        DXTRow["Priority"] = Convert.ToString(jobj["issues"][i]["fields"]["priority"]["name"]);
                        DDa.Priority = Convert.ToString(jobj["issues"][i]["fields"]["priority"]["name"]);
                        DXTRow["Severity"] = "";
                        DDa.Severity = "";
                        DDa.Creation_Date = Convert.ToString(jobj["issues"][i]["fields"]["created"]).Split(' ')[0];
                        try
                        {
                            DXTRow["Severity"] = Convert.ToString(jobj["issues"][i]["fields"]["customfield_10344"]["value"]);
                            DDa.Severity = Convert.ToString(jobj["issues"][i]["fields"]["customfield_10344"]["value"]);
                        }
                        catch
                        {
                        }
                        try
                        {
                            DXTRow["Resolution"] = Convert.ToString(jobj["issues"][i]["fields"]["resolution"]["name"]);
                            DDa.Resolution = Convert.ToString(jobj["issues"][i]["fields"]["resolution"]["name"]);
                        }
                        catch
                        {
                            DXTRow["Resolution"] = "";
                            DDa.Resolution = "";
                        }
                        DXT.Rows.Add(DXTRow);
                        DefectArray.Add(DDa);
                        j = j + 1;
                        if (j == refint)
                        {
                            float per = ((((float)i + 1) / (float)jsize)) * 100;
                            j = 0;

                            Console.WriteLine("Percentage Complete: " + per.ToString("0.00"));
                        }
                    }
                    if(jsize <100)
                    {
                        break;
                    }
                    stindx = stindx + 100;
                }
                DataTable SummarySprintStatus = SortDatatable_H_V_Count("Sprint", "Status", DXT);
                DataTable SummarySprintPriority = SortDatatable_H_V_Count("Sprint", "Priority", DXT);
                DataTable SummarySprintSeverity = SortDatatable_H_V_Count("Sprint", "Severity", DXT);
                DataTable SummaryDefect = SortDatatable_H_Count("Status", DXT);
                DataTable SummaryDefectRolling = SortDatatable_H_RollingCount("Creation Date", DXT);
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Sheets[1];
                worksheet.Name = "Dafect Data";
                Excel.Worksheet worksheet2 = workbook.Sheets.Add();
                worksheet2.Name = "Summary Data";
                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                var columns = DXT.Columns.Count;
                var rows = DXT.Rows.Count;

                Excel.Range range = worksheet.Range["A1", String.Format("{0}{1}", GetExcelColumnName(columns), rows + 1)];

                object[,] data = new object[rows + 1, columns];

                for (int rowNumber = 0; rowNumber < rows; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                        {
                            data[rowNumber, columnNumber] = (DXT.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                    {
                        data[rowNumber + 1, columnNumber] = Convert.ToString(DXT.Rows[rowNumber][columnNumber]);
                    }
                }
                Console.WriteLine("....Range.....");
                range.Value = data;
                Console.WriteLine("....Summary.....");
                var columns2 = SummarySprintStatus.Columns.Count;
                var rows2 = SummarySprintStatus.Rows.Count;
                object[,] data2 = new object[rows2 + 1, columns2];
                for (int rowNumber = 0; rowNumber < rows2; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns2; columnNumber++)
                        {
                            data2[rowNumber, columnNumber] = (SummarySprintStatus.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns2; columnNumber++)
                    {
                        data2[rowNumber + 1, columnNumber] = Convert.ToString(SummarySprintStatus.Rows[rowNumber][columnNumber]);
                    }
                }
                worksheet2.Activate();
                Excel.Range range2 = worksheet2.Range["A1", String.Format("{0}{1}", GetExcelColumnName(columns2), rows2 + 1)];
                range2.Value = data2;
                var columns3 = SummarySprintPriority.Columns.Count;
                var rows3 = SummarySprintPriority.Rows.Count;
                object[,] data3 = new object[rows3 + 1, columns3];
                Console.WriteLine("....Summary Priority..... " + columns3.ToString() + " .... " + rows3.ToString());
                for (int rowNumber = 0; rowNumber < rows3; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns3; columnNumber++)
                        {
                            data3[rowNumber, columnNumber] = (SummarySprintPriority.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns3; columnNumber++)
                    {
                        data3[rowNumber + 1, columnNumber] = Convert.ToString(SummarySprintPriority.Rows[rowNumber][columnNumber]);
                    }
                }
                worksheet2.Activate();
                Excel.Range range3 = worksheet2.Range["J1", String.Format("{0}{1}", GetExcelColumnName(columns3 + 9), rows3 + 1)];
                range3.Value = data3;
                var columns4 = SummarySprintSeverity.Columns.Count;
                var rows4 = SummarySprintSeverity.Rows.Count;
                object[,] data4 = new object[rows4 + 1, columns4];
                for (int rowNumber = 0; rowNumber < rows4; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns4; columnNumber++)
                        {
                            data4[rowNumber, columnNumber] = (SummarySprintSeverity.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns4; columnNumber++)
                    {
                        data4[rowNumber + 1, columnNumber] = Convert.ToString(SummarySprintSeverity.Rows[rowNumber][columnNumber]);
                    }
                }
                worksheet2.Activate();
                Excel.Range range4 = worksheet2.Range["S1", String.Format("{0}{1}", GetExcelColumnName(columns4 + 18), rows4 + 1)];
                range4.Value = data4;
                var columns5 = SummaryDefect.Columns.Count;
                var rows5 = SummaryDefect.Rows.Count;
                Console.WriteLine("SummaryDefect Rows Count: " + rows5.ToString());
                object[,] data5 = new object[rows5 + 1, columns5];
                for (int rowNumber = 0; rowNumber < rows5; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns5; columnNumber++)
                        {
                            data5[rowNumber, columnNumber] = (SummaryDefect.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns5; columnNumber++)
                    {
                        data5[rowNumber + 1, columnNumber] = Convert.ToString(SummaryDefect.Rows[rowNumber][columnNumber]);
                    }
                }
                worksheet2.Activate();
                Excel.Range range5 = worksheet2.Range["AA1", String.Format("{0}{1}", GetExcelColumnName(columns5 + 26), rows5 + 1)];
                range5.Value = data5;
                var columns6 = SummaryDefectRolling.Columns.Count;
                var rows6 = SummaryDefectRolling.Rows.Count;
                object[,] data6 = new object[rows6 + 1, columns6];
                for (int rowNumber = 0; rowNumber < rows6; rowNumber++)
                {
                    if (rowNumber == 0)
                    {
                        for (int columnNumber = 0; columnNumber < columns6; columnNumber++)
                        {
                            data6[rowNumber, columnNumber] = (SummaryDefectRolling.Columns[columnNumber].ColumnName);
                        }
                    }
                    for (int columnNumber = 0; columnNumber < columns6; columnNumber++)
                    {
                        data6[rowNumber + 1, columnNumber] = Convert.ToString(SummaryDefectRolling.Rows[rowNumber][columnNumber]);
                    }
                }
                worksheet2.Activate();
                Excel.Range range6 = worksheet2.Range["AE1", String.Format("{0}{1}", GetExcelColumnName(columns6 + 30), rows6 + 1)];
                range6.Value = data6;                
                Console.WriteLine("Saving Excel Workbook");
                if (File.Exists(@"C:\km\DefectDetails_" + Prj + ".xlsx"))
                {
                    File.Delete(@"C:\km\DefectDetails_" + Prj + ".xlsx");
                }
                workbook.SaveAs(@"C:\km\DefectDetails_" + Prj + ".xlsx");
                workbook.Close();
                Marshal.ReleaseComObject(application);                
                Console.WriteLine("Done....");
                var prettyJson = System.Text.Json.JsonSerializer.Serialize(DefectArray, new JsonSerializerOptions { WriteIndented = true });
                DataSet SPLUNKSEND = new DataSet();
                DXT.TableName = "DEFECTDATA";
                SPLUNKSEND.DataSetName = "SPLUNKSEND";
                SPLUNKSEND.Tables.Add(DXT);
                File.WriteAllText(@"c:\km\" + Prj+ "Defectdata.json", prettyJson);
                UploadToSPlunk(prettyJson);
                SPLUNKSEND.WriteXml(@"c:\km\DefectDataXML.xml");
                
            }
            catch (Exception E)
            {
                Console.WriteLine("Error In The Code: " + E.Message);
            }
        }
        public static void UploadToSPlunk(string REQBODY)
        {
            try
            {
                string _URL = "https://http-inputs.fisherinvestments.splunkcloud.com/services/collector/raw";
                Console.WriteLine(REQBODY);
                //Create Request
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                //Request Attributes
                request.Method = WebRequestMethods.Http.Post;
                request.Accept = "application/json";
                request.ContentType = "application/json";
                request.UserAgent = "AEQE Script";
                //Request Headers
                request.Headers.Add("Authorization", GetAuth("SPLUNKAUTH"));
                byte[] data = System.Text.Encoding.ASCII.GetBytes(REQBODY);
                request.ContentLength = data.Length;
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(data, 0, data.Length);
                requestStream.Close();
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                Console.WriteLine(responseString);
                request.Abort();
                response.Close();
            }
            catch (Exception E)
            {
                Console.WriteLine("Error In Pushing To Splunk: " + E.Message);
            }
        }
        public static System.Data.DataTable SortDatatable_H_V_Count(string Horizontal,string Vertical,System.Data.DataTable DT)
        {
            try
            {
                System.Data.DataTable DXR = new System.Data.DataTable();
                DataView view = new DataView(DT);
                DataTable distinctValuesH = view.ToTable(true, Horizontal);
                DataView view1 = new DataView(DT);
                DataTable distinctValuesV = view1.ToTable(true, Vertical);
                int vcount = distinctValuesV.Rows.Count;
                DXR.Columns.Add(Horizontal);
                int dCount = DT.Rows.Count;
                
                for (int x=0;x<vcount;x++)
                {
                    DXR.Columns.Add(Convert.ToString(distinctValuesV.Rows[x][0]));
                }
                int colcount = DXR.Columns.Count;
                int hcount = distinctValuesH.Rows.Count;
                for (int x = 0; x < hcount; x++)
                {
                    DataRow DXRROW = DXR.NewRow();
                    string horizontalval  = Convert.ToString(distinctValuesH.Rows[x][0]);
                    DXRROW[0] = horizontalval;
                    for (int w = 1; w < colcount; w++)
                    {
                        string colval = DXR.Columns[w].ColumnName;
                        int count = 0;
                        for (int y = 0; y < dCount; y++)
                        {
                            if (Convert.ToString(DT.Rows[y][Horizontal]) == horizontalval && Convert.ToString(DT.Rows[y][Vertical]) == colval)
                            {
                                count++;
                            }
                        }
                        DXRROW[w] = count;
                    }
                    DXR.Rows.Add(DXRROW);
                }
                return DXR;
            }
            catch
            {
                return new System.Data.DataTable();
            }
        }
        public static System.Data.DataTable SortDatatable_H_Count(string Horizontal, System.Data.DataTable DT)
        {
            try
            {
                System.Data.DataTable DXR = new System.Data.DataTable();
                DataView view = new DataView(DT);
                DataTable distinctValuesH = view.ToTable(true, Horizontal);
                DXR.Columns.Add(Horizontal);
                DXR.Columns.Add("Count");
                int dCount = DT.Rows.Count;
                int colcount = DXR.Columns.Count;
                int hcount = distinctValuesH.Rows.Count;
                for (int x = 0; x < hcount; x++)
                {
                    DataRow DXRROW = DXR.NewRow();
                    string horizontalval = Convert.ToString(distinctValuesH.Rows[x][0]);
                    DXRROW[0] = horizontalval;
                    for (int w = 1; w < colcount; w++)
                    {
                        int count = 0;
                        for (int y = 0; y < dCount; y++)
                        {
                            if (Convert.ToString(DT.Rows[y][Horizontal]) == horizontalval)
                            {
                                count++;
                            }
                        }
                        DXRROW[w] = count;
                    }
                    DXR.Rows.Add(DXRROW);
                }
                return DXR;
            }
            catch
            {
                return new System.Data.DataTable();
            }
        }
        public static System.Data.DataTable SortDatatable_H_RollingCount(string Horizontal, System.Data.DataTable DT)
        {
            try
            {
                System.Data.DataTable DXR = new System.Data.DataTable();
                DataView view = new DataView(DT);
                DataTable distinctValuesH = view.ToTable(true, Horizontal);

                DXR.Columns.Add(Horizontal);
                DXR.Columns.Add("Count");

                int dCount = DT.Rows.Count;

                int colcount = DXR.Columns.Count;
                int hcount = distinctValuesH.Rows.Count;
                int count = 0;
                for (int x = 0; x < hcount; x++)
                {
                    DataRow DXRROW = DXR.NewRow();
                    string horizontalval = Convert.ToString(distinctValuesH.Rows[x][0]);
                    DXRROW[0] = horizontalval;
                    for (int w = 1; w < colcount; w++)
                    {                        
                        for (int y = 0; y < dCount; y++)
                        {
                            if (Convert.ToString(DT.Rows[y][Horizontal]) == horizontalval)
                            {
                                count++;
                            }
                        }
                        DXRROW[w] = count;
                    }
                    DXR.Rows.Add(DXRROW);
                }
                return DXR;
            }
            catch
            {
                return new System.Data.DataTable();
            }
        }
        
        private static List<string> GetSprint(string defectKey)
        {
            List<string> sprintdata = new();
            try
            {                
                string _URL = "https://fish-net.atlassian.net/rest/agile/1.0/issue/" + defectKey;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                request.Method = WebRequestMethods.Http.Get;
                request.Accept = "appliation/json";
                request.ContentType = "application/json";
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0";
                request.Headers.Add("Authorization"
                                  , GetAuth("JIRAAUTH"));
                request.Headers.Add("Accept", "application/json");
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                //json parsing
                JObject jobj = JObject.Parse(responseString);
                response.Close();
                request.Abort();
                sprintdata.Add(Convert.ToString(jobj["fields"]["customfield_10020"][0]["name"]));
                sprintdata.Add(Convert.ToString(jobj["fields"]["customfield_10020"][0]["startDate"]));
                sprintdata.Add(Convert.ToString(jobj["fields"]["customfield_10020"][0]["endDate"]));
                return sprintdata;
            }
            catch(Exception )
            {
                sprintdata.Add("");
                sprintdata.Add("");
                sprintdata.Add("");
                return sprintdata;
            }
        }
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
 
        public static string GetAuth(string refstr = "JIRAAUTH")
        {
            try
            {
                string fcontent = File.ReadAllText(@"C:\TestFramework\Authdata.json");
                JObject jobj = JObject.Parse(fcontent);
                switch (refstr)
                {
                    case "JIRAAUTH":
                        return jobj["JIRAAUTH"].ToString();
                    case "ZEPHYRAUTH":
                        return jobj["ZEPHYRAUTH"].ToString();
                    case "SPLUNKAUTH":
                        return jobj["SPLUNKAUTH"].ToString();
                }
                return "";
            }
            catch
            {
                return "";
            }
        }
    }
}