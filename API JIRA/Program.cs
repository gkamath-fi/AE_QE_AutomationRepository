using System.Data;
using System.Text.RegularExpressions;
using System.Net;
using Newtonsoft.Json.Linq;
using DataTable = System.Data.DataTable;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Text;
using System.Web;
using Newtonsoft.Json;

//Changes Pushed by gkamath on 3/17/2025
namespace API //Changes to API


namespace API

{    
    [DisplayName("Jira Defect Data")]
    public class DefectData
    {
        [DisplayName("Project")]
        public string Project { get; set; } = "";
        [DisplayName("Creation Date")]
        public string Creation_Date { get; set; } = "";
        [DisplayName("Business Unit")]
        public string BU{ get; set; } = "";
        [DisplayName("Defect Key")]
        public string Key { get; set; } = "";
        [DisplayName("Defect ID")]
        public string ID { get; set; } = "";
        [DisplayName("Defect Age")]
        public string DefectAge { get; set; } = "";
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
        public string Type { get; set; } = "Bug";
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

        //Automation reporting for PCG new change on 3/10/2025

        [DisplayName("Linked Test")]
        public string LinkedTest { get; set; } = "";
    }  

    public class TEMPLINKVAL
    {
        public string url { get; set; } = "";
        public string value { get; set; } = "";
        
    }
    public class PRJDATA
    {
        public string PrName { get; set; } = "";
        public string BuName { get; set; } = "";
        public PRJDATA(string pn,string bn)
        {
            PrName = pn;
            BuName = bn;
        }
    }
    public static class JIRAALLDefects
    {
        public static List<TEMPLINKVAL> TLIST = new();
        private static BackgroundWorker fetchworker = new();
        public static List<DefectData> DefectArray = new List<DefectData>();
        public static List<TestCaseData> TCArray = new List<TestCaseData>();
        static int threadcount = 0;
        public static void  Main()
        {
            try
            {
                Console.WriteLine("Enter Option J(JIRA and Zehyr report) N: Neoload Validation: ");
                string option = Console.ReadLine();
                switch (option)
                {
                    case "J":
                        int BufferSize = 128;
                        Console.WriteLine("Time Start: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        string[] files = Directory.GetFiles(@"c:\TestFramework\", "Projects*.txt");
                        foreach (string fname in files)
                        {
                            BackgroundWorker worker = new BackgroundWorker { WorkerReportsProgress = true, WorkerSupportsCancellation = true };
                            worker.DoWork += worker_DoWork;
                            worker.WorkerReportsProgress = true;
                            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
                            worker.RunWorkerAsync(fname);
                            threadcount++;
                        }
                        while (threadcount > 0)
                        {
                            System.Threading.Thread.Sleep(50000);
                        }
                        Console.WriteLine("Time End: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        break;
                    case "N":
                        List<string> SERVERLIST = GETNEOLOADSERVERS();
                        NEOLOADDATA(SERVERLIST);
                        break;
                }
            }
            catch
            {
            }
        }
        
        private static void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                int BufferSize = 120;
                string fname = (string)(e.Argument);
                string PrjName = "";
                string Bunit = "";
                using (var fileStream = File.OpenRead(fname))
                {
                    using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                    {
                        String line;
                        while ((line = streamReader.ReadLine()) != null)
                        {
                            if (!string.IsNullOrEmpty(line.Trim()))
                            {
                                PrjName = line.Split(',')[0].Trim();
                                Bunit = line.Split(',')[1].Trim();
                              
                                JIRADEFECTdata(PrjName,Bunit).Wait();
                            
                                GETZEPHYRSCALEDATATESTCASE(PrjName, Bunit);
                                GETZEPHYRSCALEDATATESTEXECUTION(PrjName, Bunit);
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }
        public static  void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                Console.WriteLine("Thread Completed");
                threadcount--;
            }
            catch
            {
            }
        }
        public class TestCaseData
        {
            [DisplayName("Business Unit")]
            public string BU { get; set; } = "";
            [DisplayName("Test Case")]
            public string TestCase { get; set; } = "";
            [DisplayName("Project")]
            public string Project { get; set; } = "";
            [DisplayName("Created On")]
            public string CreatedOn { get; set; } = "";
            public string Automated { get; set; } = "No";
            public string Purpose { get; set; } = "";
        }
        public class TextExecutionData
        {
            [DisplayName("Business Unit")]
            public string BU { get; set; } = "";
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
        public static void TCAutomated(string TCIDS)
        {
            try
            {
                string TCID = TCIDS.Split("/")[5];
                int DSize = TCArray.Count;
              
                for (int j = 0; j < DSize; j++)
                {
                    if (TCArray[j].TestCase == TCID)
                    {
                        TCArray[j].Automated = "Yes";
                        break;
                    }
                }
                
            }
            catch
            {

            }
        }
        public static void GETISSUELINK(string testexecid)
        {
            try
            {
                string ZephyrAuth = GetAuth("ZEPHYRAUTH");
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://api.zephyrscale.smartbear.com/v2/testexecutions/" + testexecid + "/links");
                request.Method = "GET";
                request.UserAgent = "AEQE Analysis";
                request.Headers.Add("Authorization", ZephyrAuth);
                request.Headers.Add("Accept", "application/json");
                request.Headers.Add("Content-Type", "application/json");
                request.KeepAlive = true;
                var response = (HttpWebResponse)request.GetResponse();
                string  responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                JObject jobj = JObject.Parse(responseString);
                JArray jArray = (JArray)jobj["issues"];
                int jsize = jArray.Count;
                int DSize = DefectArray.Count;
                for (int i = 0; i < jsize; i++)
                {
                    string issueid = Convert.ToString(jobj[i]["issueId"]);
                    for(int j=0;j<DSize;j++)
                    {
                        if (DefectArray[j].ID == issueid)
                        {
                            DefectArray[j].LinkedTest = testexecid;
                            break;
                        }
                    }
                }
            }
            catch
            {

            }
        }
        public static void GETZEPHYRSCALEDATATESTEXECUTION(string PrjName,string  Bunit)
        {
            try
            {
                int startindx = 0;
                int maxresults = 50;
                bool continueloop = true;                
                string Prj = PrjName;
                string ZephyrAuth = GetAuth("ZEPHYRAUTH");
                string JIRAAUTH = GetAuth("JIRAAUTH");
                List<TextExecutionData> TEArray = new List<TextExecutionData>();
                string purpose = Prj + " Test Execution Data";
                string autostatus = "";
                int jsize = 0;
                string responseString;
                while (continueloop)
                {
                    
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://api.zephyrscale.smartbear.com/v2/testexecutions?projectKey=" + Prj + "&maxResults=" + maxresults.ToString() + "&startAt=" + startindx.ToString());
                    request.Method = "GET";
                    request.UserAgent = "AEQE Analysis";
                    request.Headers.Add("Authorization", ZephyrAuth);
                    request.Headers.Add("Accept", "application/json");
                    request.Headers.Add("Content-Type", "application/json");
                    request.KeepAlive = true;
                    var response = (HttpWebResponse)request.GetResponse();
                    responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                    JObject jobj = JObject.Parse(responseString);
                    JArray jArray = (JArray)jobj["values"];
                    jsize = jArray.Count;
                    for (int i = 0; i < jsize; i++)
                    {
                        TextExecutionData teitem = new TextExecutionData();
                        teitem.Purpose = purpose;
                        
                        teitem.TestExecutionKey = jobj["values"][i]["key"].ToString();
                        
                        try
                        {
                            teitem.BU = Bunit;
                            teitem.Project = GetProjectData(ZephyrAuth, JIRAAUTH, jobj["values"][i]["project"]["self"].ToString(), jobj["values"][i]["project"]["id"].ToString());
                        }
                        catch
                        {
                        }                       
                        teitem.TestCase = GetProjectData(ZephyrAuth, JIRAAUTH, jobj["values"][i]["testCase"]["self"].ToString(), jobj["values"][i]["testCase"]["id"].ToString(), "Test Case");                      
                        teitem.Status = GetProjectData(ZephyrAuth, JIRAAUTH, jobj["values"][i]["testExecutionStatus"]["self"].ToString(), jobj["values"][i]["testExecutionStatus"]["id"].ToString(), "Status");
                        try
                        {
                            teitem.Environment = GetProjectData(ZephyrAuth, JIRAAUTH, jobj["values"][i]["environment"]["self"].ToString(), jobj["values"][i]["environment"]["id"].ToString(), "Environment");
                            if(string.IsNullOrEmpty(teitem.Environment))
                            {
                                teitem.Environment = "Not Mentioned";
                            }
                        }
                        catch
                        {
                            
                            teitem.Environment = "Not Mentioned";
                        }
                        teitem.ActualEndDate = jobj["values"][i]["actualEndDate"].ToString().Split(' ')[0];
                        teitem._time = jobj["values"][i]["actualEndDate"].ToString().Split(' ')[0];
                        autostatus = "Automated";
                        if(jobj["values"][i]["automated"].ToString()=="False")
                        {
                            autostatus = "Manual";
                        }
                        else
                        {
                            TCAutomated(jobj["values"][i]["testCase"]["self"].ToString());
                            //if (teitem.Status == "Blocked" || teitem.Status == "Fail")
                            //{
                            //    GETISSUELINK(teitem.TestExecutionKey);
                            //}
                        }
                        teitem.ExecutionType= autostatus;
                        try
                        {
                            if (!string.IsNullOrEmpty(jobj["values"][i]["assignedToId"].ToString()))
                            {
                                teitem.AssignedTo = GetProjectData(ZephyrAuth, JIRAAUTH, "https://fish-net.atlassian.net/rest/api/3/user?accountId=" + jobj["values"][i]["assignedToId"].ToString(), jobj["values"][i]["assignedToId"].ToString(), "Assigned To");
                            }
                            if (string.IsNullOrEmpty(teitem.AssignedTo))
                            {
                                teitem.AssignedTo = "Unassigned";
                            }
                        }
                        catch
                        {
                            teitem.AssignedTo = "Unassigned";
                        }
                        try
                        {
                            teitem.TestCycle = GetProjectData(ZephyrAuth,JIRAAUTH,jobj["values"][i]["testCycle"]["self"].ToString(), jobj["values"][i]["environment"]["id"].ToString(), "Test Cycle");
                        }
                        catch
                        {
                            teitem.TestCycle = "Not Specified";
                        }
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
                var prettyJson = System.Text.Json.JsonSerializer.Serialize(TEArray, new JsonSerializerOptions { WriteIndented = true });
                UploadToSPlunk(prettyJson);
                //var prettyJson1 = System.Text.Json.JsonSerializer.Serialize(DefectArray, new JsonSerializerOptions { WriteIndented = true });
                //UploadToSPlunk(prettyJson1); 
                prettyJson = System.Text.Json.JsonSerializer.Serialize(TCArray, new JsonSerializerOptions { WriteIndented = true });
                UploadToSPlunk(prettyJson);

                TEArray = new();
                TCArray = new();
            }
            catch (Exception E)
            {
                MessageBox.Show(E.Message, "Error");
            }
        }
        public static void GETZEPHYRSCALEDATATESTCASE(string PrjName, string Bunit)
        {
            try
            {
                int startindx = 0;
                int maxresults = 50;
                string Prj = PrjName;
                bool continueloop = true;
                string ZephyrAuth = GetAuth("ZEPHYRAUTH");
                string JIRAAUTH = GetAuth("JIRAAUTH");
                
                string Purpose = Prj + " Test Case Data";
                string _URL = "";
                int jsize = 0;
                int i;
                while (continueloop)
                {
                    _URL = "https://api.zephyrscale.smartbear.com/v2/testcases?projectKey=" + Prj + "&maxResults=" + maxresults.ToString() + "&startAt=" + startindx.ToString();
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                    request.Method = "GET";
                    request.UserAgent = "AEQE Analysis";
                    request.Headers.Add("Authorization", ZephyrAuth);
                    request.Headers.Add("Accept", "application/json");
                    request.Headers.Add("Content-Type", "application/json");
                    request.KeepAlive = true;
                    var response = (HttpWebResponse)request.GetResponse();
                    var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                    request.Abort();
                    response.Close();
                    JObject jobj = JObject.Parse(responseString);
                    JArray jArray = (JArray)jobj["values"];
                    jsize = jArray.Count;
                    for (i = 0; i < jsize; i++)
                    {
                        TestCaseData TCD = new();
                        TCD.TestCase = jobj["values"][i]["key"].ToString();
                        TCD.Purpose = Purpose;                        
                        try
                        {
                            TCD.BU = Bunit;
                            TCD.Project = GetProjectData(ZephyrAuth, JIRAAUTH,jobj["values"][i]["project"]["self"].ToString(),jobj["values"][i]["project"]["id"].ToString());
                        }
                        catch
                        {
                        }                        
                        TCD.CreatedOn = jobj["values"][i]["createdOn"].ToString().Split(' ')[0];
                        TCArray.Add(TCD);
                    }
                    if (jsize <50)
                    {
                        continueloop = false;
                    }
                    startindx = startindx + 50;
                } 
                //var prettyJson = System.Text.Json.JsonSerializer.Serialize(TCArray, new JsonSerializerOptions { WriteIndented = true });
               // UploadToSPlunk(prettyJson);
            }
            catch(Exception E)
            {
                MessageBox.Show(E.Message);
            }
        }
        public static string GetProjectData(string ZAUTH, string JAUTH,string _URL,string prjID,string refstr="Project")
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
                    request.Headers.Add("Authorization", ZAUTH);
                }
                else
                {
                    request.Headers.Add("Authorization", JAUTH);
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
                request.Headers.Add("Authorization", GetAuth("JIRAAUTH"));
                request.Headers.Add("Accept", "application/json");
                request.Headers.Add("Content-Type", "application/json");
                request.KeepAlive = true;
                var response = (HttpWebResponse)request.GetResponse();
                
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                
                string rx = @"([A-Za-z]{2})-([A-Za-z]{2})-([A-Za-z]{1})(\d{5}):7100";
                request.Abort();
                response.Close();
                MatchCollection m = Regex.Matches(responseString, rx, RegexOptions.IgnoreCase);
                foreach (Match match in m)
                {
                    Console.WriteLine("Server name : {0}, Index : {1}", match.Value, match.Index);
                    if (match.Value.ToString().ToUpper().StartsWith("CP"))
                    {
                        Prodservers.Add(match.Value.ToString().ToUpper());
                    }
                }
            }
            catch (Exception E)
            {
                MessageBox.Show(E.Message, "Error");
            }
            return Prodservers;
        }
        public static void NEOLOADDATA(List<string> SERVERLIST)
        {
            try
            {
                string _URL = "https://neoload-api.saas.neotys.com/v3/resources/zones";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                request.Method = WebRequestMethods.Http.Get;
                request.Accept = "appliation/json";
                request.ContentType = "application/json";
                request.UserAgent = "AEQE Analysis";
                request.Headers.Add("accountToken", GetAuth("NEOLOADAUTH"));
                request.Headers.Add("Accept", "application/json");
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                JToken parsedJson = JToken.Parse(responseString);
                var beautified = parsedJson.ToString(Formatting.Indented);
                var minified = parsedJson.ToString(Formatting.None);

                File.WriteAllText(@"c:\Km\NeoloadData.json", beautified);
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
                for (int x = 0; x < jsize; x++)
                {
                    DataRow DR = DTAB.NewRow();
                    DR["Resource Id"] = (jArray[0]["id"].ToString());
                    DR["Resource Name"] = (jArray[0]["name"].ToString());
                    DR["Resource Type"] = (jArray[0]["type"].ToString());
                    int controllerCount = jArray[x]["controllers"].Count();
                    DR["Load Controllers"] = "";
                    for (int y = 0; y < controllerCount; y++)
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
                        if (SERVERLIST.Contains(jArray[x]["loadgenerators"][y]["name"].ToString().ToUpper()))
                        {
                            Console.WriteLine("Validated Server: " + jArray[x]["loadgenerators"][y]["name"].ToString().ToUpper());
                            SERVERLIST.Remove(jArray[x]["loadgenerators"][y]["name"].ToString().ToUpper());
                        }
                    }
                    DTAB.Rows.Add(DR);
                }
                if (SERVERLIST.Count > 0)
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
                range.Value = data;
                if (File.Exists(@"C:\km\NeoloadResources" + ".xlsx"))
                {
                    File.Delete(@"C:\km\NeoloadResources" + ".xlsx");
                }
                workbook.SaveAs(@"C:\km\NeoloadResources" + ".xlsx");
                workbook.Close();
            }
            catch (Exception E)
            {
                MessageBox.Show(E.Message);
            }
        }
        public static async Task JIRADEFECTdata(string PrjName,string Bunit)
        {
            try
            {
                string Prj = PrjName;
                
                string AUTHSTR = GetAuth("JIRAAUTH");
                string Purpose = Prj + " Defect Reporting"; 
                int stindx = 0;
                var method = HttpMethod.Get;
                DateTime CDate = DateTime.Now;
                int month, date, year;
                string creationdate = "";
                int Dage = 0;
                string _URL = "";
                string responseString = "";
                int jsize = 0,i;
                JObject jobj;
                JArray jArray;
                HttpClient client;
                HttpRequestMessage request;
                while (1 == 1)
                {
                    _URL = "https://fish-net.atlassian.net/rest/api/2/search?jql=project=" + Prj + " And type=Bug order by created asc&fields=key,sprint,status,priority,issuetype,priority,assignee,summary,customfield_10344,created,resolution&startAt=" + stindx.ToString() + "&maxResults=100";
                    client = new HttpClient();
                    request = new HttpRequestMessage();                   
                    request.RequestUri = new Uri(_URL);
                    request.Method =method;
                    request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Add("User-Agent", "QE Analysis");
                    request.Headers.Add("Authorization", AUTHSTR);
                    var response = await client.SendAsync(request);
                     responseString = await response.Content.ReadAsStringAsync();
                    jobj = JObject.Parse(responseString);
                    jArray = (JArray)jobj["issues"];
                    jsize = jArray.Count;                  
                    Dage = 0;
                    for (i = 0; i < jsize; i++)
                    {
                        DefectData DDa = new DefectData();
                        DDa.ID= jobj["issues"][i]["id"].ToString();
                        DDa.Key = jobj["issues"][i]["key"].ToString();
                        DDa.Project = Prj;
                        DDa.BU = Bunit;
                        DDa.Status = jobj["issues"][i]["fields"]["status"]["name"].ToString();
                        DDa.Summary = jobj["issues"][i]["fields"]["summary"].ToString();                   
                        DDa.Purpose = Purpose;
                        List<string> sdata = GetSprint(jobj["issues"][i]["key"].ToString());
                        if (!string.IsNullOrEmpty(sdata[0]))
                        {
                            DDa.Sprint = sdata[0];
                        }
                        else
                        {
                            DDa.Sprint = "Not Defined";
                        }
                        DDa.Sprint_Start_Date = sdata[1];
                        DDa.Issue_Type = jobj["issues"][i]["fields"]["issuetype"]["name"].ToString();
                        DDa.Assignee = Convert.ToString(jobj["issues"][i]["fields"]["Assignee"]);
                        if (!string.IsNullOrEmpty(Convert.ToString(jobj["issues"][i]["fields"]["priority"]["name"])))
                        {
                            DDa.Priority = Convert.ToString(jobj["issues"][i]["fields"]["priority"]["name"]);
                        }
                        else
                        {
                            DDa.Priority = "Not Defined";
                        }
                        creationdate = Convert.ToString(jobj["issues"][i]["fields"]["created"]).Split(' ')[0];
                        string[] datemon = creationdate.Split('/');
                        int.TryParse(datemon[0], out month);
                        int.TryParse(datemon[1], out date);
                        int.TryParse(datemon[2], out year);
                         CDate = new DateTime(year,month,date);
                         Dage = (DateTime.Now - CDate).Days;
                        if(Dage<=10)
                        {
                            DDa.DefectAge = "<=10 Days";
                        }
                        else if (Dage <= 30)
                        {
                            DDa.DefectAge = "<=30 Days";
                        }
                        else if (Dage <= 60)
                        {
                            DDa.DefectAge = "<=60 Days";
                        }
                        else if (Dage <= 100)
                        {
                            DDa.DefectAge = "<=100 Days";
                        }
                        else
                        {
                            DDa.DefectAge = ">100 Days";
                        }
                        DDa.Creation_Date = creationdate;

                        try
                        {
                            DDa.Severity = Convert.ToString(jobj["issues"][i]["fields"]["customfield_10344"]["value"]);
                        }
                        catch
                        {
                            DDa.Severity = "Not Defined";
                        }
                        try
                        {
                            DDa.Resolution = Convert.ToString(jobj["issues"][i]["fields"]["resolution"]["name"]);
                        }
                        catch
                        {
                            DDa.Resolution = "Not Defined";
                        }
                        DefectArray.Add(DDa);
                    }
                    if (jsize < 100)
                    {
                        break;
                    }
                    stindx = stindx + 100;
                }
                var prettyJson = System.Text.Json.JsonSerializer.Serialize(DefectArray, new JsonSerializerOptions { WriteIndented = true });
                UploadToSPlunk(prettyJson);
                DefectArray = new();
            }
            catch(Exception E)
            {
                Console.WriteLine("Error: " + E.Message);
            }
        }
        public static void UploadToSPlunk(string REQBODY)
        {
            try
            {
                string _URL = "https://http-inputs.fisherinvestments.splunkcloud.com/services/collector/raw";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_URL);
                request.Method = WebRequestMethods.Http.Post;
                request.Accept = "application/json";
                request.ContentType = "application/json";
                request.UserAgent = "AEQE Script";
                request.Headers.Add("Authorization", GetAuth("SPLUNKAUTH"));
                byte[] data = System.Text.Encoding.ASCII.GetBytes(REQBODY);
                request.ContentLength = data.Length;
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(data, 0, data.Length);
                requestStream.Close();
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                request.Abort();
                response.Close();
            }
            catch (Exception E)
            {
                Console.WriteLine("Error In Pushing To Splunk: " + E.Message);
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
                request.UserAgent = "QE Analysis";
                request.Headers.Add("Authorization" , GetAuth("JIRAAUTH"));
                request.Headers.Add("Accept", "application/json");
                var response = (HttpWebResponse)request.GetResponse();
                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                JObject jobj = JObject.Parse(responseString);
                response.Close();
                request.Abort();
                sprintdata.Add(Convert.ToString(jobj["fields"]["customfield_10020"][0]["name"]));
                sprintdata.Add(Convert.ToString(jobj["fields"]["customfield_10020"][0]["startDate"]));
                return sprintdata;
            }
            catch(Exception )
            {
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
                    case "NEOLOADAUTH":
                        return jobj["NEOLOADAUTH"].ToString();
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
