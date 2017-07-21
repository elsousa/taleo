using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;
using RestSharp;
using RestSharp.Serializers;
using System.Runtime.Serialization;
using RestSharp.Deserializers;
using System.Web.Script.Serialization;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace TesteTaleo
{
    class Program
    {
        /// <summary>
        /// Install-Package Newtonsoft.Json -Version 8.0.3 
        /// </summary>
        /// <param name="args"></param>

        static string _token = string.Empty;

        static void Main(string[] args)
        {
            string execucao = string.Empty;

            //ConsultarAccount();
            while (execucao.ToLower() != "stop")
            {
                execucao = Console.ReadLine();

                //LerArquviXml();

                Console.WriteLine("Valor selecionado..: {0}", execucao);

                switch (execucao)
                {
                    case "1":
                        ConsultaToken();
                        break;
                    case "2":
                        ConsultarAccountById("45");
                        break;
                    case "3":
                        //InserirEmployeeGoal();
                        LerArquvoXml();
                        break;
                    default:
                        ConsultaLogout();
                        break;
                }
            }
        }

        public static void ConsultarAccount()
        {
            string token = string.Empty;

            //token = ConsultaToken();

            //ConsultarAccountById(token, "45");
        }

        private static void ConsultarAccountById(string id)
        {
            WebProxy proxy = ConfiguracaoProxy();
            IRestClient restClient = new RestClient("https://localhost/object/account/45");
            restClient.Proxy = proxy;

            var request = new RestRequest(Method.GET);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("authToken", _token);
            request.RootElement = "account";

            var json = restClient.Execute(request).Content;

            JObject objParserd = JObject.Parse(json);

            JObject arrayObject1 = (JObject)objParserd["response"]["account"];
            var myOutput = JsonConvert.DeserializeObject<Account>(arrayObject1.ToString());
        }

        private static WebProxy ConfiguracaoProxy()
        {
            WebProxy proxy = new WebProxy("proxypac", 8080);
            proxy.UseDefaultCredentials = true;
            proxy.BypassProxyOnLocal = false;
            return proxy;
        }

        private static void ConsultaToken()
        {
            string retorno = string.Empty;
            string url = "https://localhost?orgCode=xcv&userName=elopes&password=xcv";
            RestClient restClient = new RestClient(url);
            WebProxy proxy = new WebProxy("proxypac", 8080);
            proxy.BypassProxyOnLocal = false;

            ICredentials credential = new NetworkCredential("eyvewp", "Abcd@5432");
            proxy.Credentials = credential;

            restClient.Proxy = proxy;

            RestRequest request = new RestRequest(Method.POST);
            IRestResponse<Token> responseToken = restClient.Execute<Token>(request);

            if (responseToken != null && responseToken.Data != null && responseToken.Data.response != null && responseToken.Data.status.success == true)
            {
                if (!string.IsNullOrEmpty(responseToken.Data.response.authToken))
                {
                    _token = responseToken.Data.response.authToken;
                }
            }
            else
            {
                var json = restClient.Execute(request).Content;

                JObject objParserd = JObject.Parse(json);

                JObject arrayObject1 = (JObject)objParserd["status"]["detail"];

                Console.WriteLine(arrayObject1.First.Last);
            }
        }

        private static void ConsultaLogout()
        {
            string retorno = string.Empty;
            string url = "https://localhost/logout";
            RestClient restClient = new RestClient(url);
            WebProxy proxy = new WebProxy("proxypac", 8080);
            proxy.BypassProxyOnLocal = false;

            ICredentials credential = new NetworkCredential("eyvewp", "Abcd@5432");
            proxy.Credentials = credential;
            restClient.Proxy = proxy;

            RestRequest request = new RestRequest(Method.POST);
            request.AddCookie("authToken", _token);

            IRestResponse responseToken = restClient.Execute(request);
        }

        private static void InserirEmployeeGoal(EmployeeGoal EmployeeGoal)
        {
            IRestResponse response = Create<EmployeeGoal>(EmployeeGoal, "/object/employeegoal");
        }

        private static EmployeeGoal TesteEmployeeGoal()
        {
            EmployeeGoalJson EmployeeGoalJson = new EmployeeGoalJson();
            EmployeeGoalJson.goalPercent = 5;
            //employeeGoal.corpGoals = "Teste Copr Goals";
            EmployeeGoalJson.createdById = 117;
            EmployeeGoalJson.creationDate = "2017-07-21T04:57PDT";
            EmployeeGoalJson.department = 57;
            EmployeeGoalJson.description = "Teste description";
            EmployeeGoalJson.division = 40;
            EmployeeGoalJson.dueDate = "2017-07-24";
            EmployeeGoalJson.employeeComment = "Comentário";
            EmployeeGoalJson.reviewManager = 55;
            EmployeeGoalJson.employeeManager = 55;
            EmployeeGoalJson.GoalAlignment = "Divisional Directive";
            EmployeeGoalJson.goalTitle = "Teste Evandro 10:53";
            EmployeeGoalJson.goalId = 181;
            EmployeeGoalJson.employeeId = 130;
            EmployeeGoalJson.employeePicture = "";
            EmployeeGoalJson.lastUpdated = "2017-07-21T04:57PDT";
            EmployeeGoalJson.location = 41;
            EmployeeGoalJson.managerComment = "";
            EmployeeGoalJson.measurement = "Teste 10:53";
            EmployeeGoalJson.goalEmplIsActive = false;
            EmployeeGoalJson.percentComplete = 35;
            EmployeeGoalJson.isPrivate = false;
            EmployeeGoalJson.region = 42;
            EmployeeGoalJson.site_teste_upit = "Teste 10:53";
            EmployeeGoalJson.status = 4;
            EmployeeGoalJson.goalApprovalStatus = 1;
            EmployeeGoalJson.reviewType = "Meio do Ano";

            EmployeeGoal EmployeeGoal = new EmployeeGoal();
            EmployeeGoal.employeegoal = EmployeeGoalJson;

            return EmployeeGoal;
        }

        private static void LerArquvoXml()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Carga_Metas.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            try
            {

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                if (i == 1)
                    continue;

                EmployeeGoalJson emp = new EmployeeGoalJson();

                emp.goalTitle = xlRange.Cells[i, 1].Value2.ToString();
                emp.reviewType = xlRange.Cells[i, 2].Value2.ToString();
                emp.description = xlRange.Cells[i, 3].Value2.ToString();
                emp.measurement = xlRange.Cells[i, 4].Value2.ToString();
                emp.GoalAlignment = xlRange.Cells[i, 5].Value2.ToString();
                emp.dueDate = ConvertDate(xlRange.Cells[i, 6].Text.ToString());
                emp.status = 4;
                emp.percentComplete = int.Parse(xlRange.Cells[i, 8].Value2.ToString());
                emp.site_teste_upit = xlRange.Cells[i, 10].Value2.ToString();
                emp.employeeManager = int.Parse(xlRange.Cells[i, 11].Value2.ToString());
                emp.goalApprovalStatus = int.Parse(xlRange.Cells[i, 12].Value2.ToString());
                emp.employeeComment = xlRange.Cells[i, 13].Value2.ToString();
                emp.creationDate = ConvertDataHora(DateTime.Now.ToShortDateString());
                emp.division = int.Parse(xlRange.Cells[i, 15].Value2.ToString());
                emp.lastUpdated = ConvertDataHora(DateTime.Now.ToShortDateString());
                emp.reviewManager = int.Parse(xlRange.Cells[i, 17].Value2.ToString());
                emp.employeePicture = string.Empty;
                emp.managerComment = string.Empty;
                emp.location = int.Parse(xlRange.Cells[i, 20].Value2.ToString());
                emp.department = int.Parse(xlRange.Cells[i, 21].Value2.ToString());
                emp.region = int.Parse(xlRange.Cells[i, 22].Value2.ToString());
                emp.createdById = int.Parse(xlRange.Cells[i, 23].Value2.ToString());
                emp.employeeId = int.Parse(xlRange.Cells[i, 24].Value2.ToString());

                EmployeeGoal empG = new EmployeeGoal()
                {
                    employeegoal = emp
                };

                InserirEmployeeGoal(empG);
            }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                ConsultaLogout();
            }
        }

        private static string ConvertDate(string data)
        {
            //2017-07-24
            var novaData = data.Split('/');
            return novaData[2] + "-" + novaData[1] + "-" + novaData[0];
        }

        //2017-07-21T04:57PDT
        private static string ConvertDataHora(string data)
        {
            var novaData = data.Split('/');
            return novaData[2] + "-" + novaData[1] + "-" + novaData[0] + "T04:57PDT";
        }

        private static IRestResponse Create<T>(object objectToUpdate, string apiEndPoint) where T : new()
        {
            var client = new RestClient(CreateBaseUrl(apiEndPoint))
            {
                Proxy = ConfiguracaoProxy()
            };

            var json = JsonConvert.SerializeObject(objectToUpdate);
            var request = new RestRequest(Method.POST);

            request.RequestFormat = DataFormat.Json;
            request.AddCookie("authToken", _token);
            request.AddParameter("application/json", json, ParameterType.RequestBody);

            var response = client.Execute<T>(request);

            return response;
        }

        private static string CreateBaseUrl(string apiEndPoint)
        {
            return "https://localhost" + apiEndPoint;
        }

    }

    public class EmployeeGoal
    {
        public EmployeeGoalJson employeegoal { get; set; }
    }

    public class EmployeeGoalJson
    {
        public int goalPercent { get; set; }
        //public string corpGoals { get; set; }
        public int createdById { get; set; }
        public string creationDate { get; set; }
        public int department { get; set; }
        public string description { get; set; }
        public int division { get; set; }
        public string dueDate { get; set; }
        public string employeeComment { get; set; }
        public int reviewManager { get; set; }
        public int employeeManager { get; set; }
        public string GoalAlignment { get; set; }
        public string goalTitle { get; set; }
        public int goalId { get; set; }
        public int employeeId { get; set; }
        public string employeePicture { get; set; }
        public string lastUpdated { get; set; }
        public int location { get; set; }
        public string managerComment { get; set; }
        public string measurement { get; set; }
        public bool goalEmplIsActive { get; set; }
        public int percentComplete { get; set; }
        public bool isPrivate { get; set; }
        public int region { get; set; }
        public string site_teste_upit { get; set; }
        public int status { get; set; }
        public int goalApprovalStatus { get; set; }
        public string reviewType { get; set; }
    }

    public class Teste
    {
        public string orgCode { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }

    public class Token
    {
        public Response response { get; set; }
        public Status status { get; set; }
    }

    public class Response
    {
        public string authToken { get; set; }
        public Account account { get; set; }
    }

    public class Status
    {
        public bool success { get; set; }
        public string detail { get; set; }
    }

    public class Account
    {
        public string name { get; set; }
        public string creationDate { get; set; }
        public string city { get; set; }
        public string description { get; set; }
        public string state { get; set; }
        public string fax { get; set; }
        public int accountId { get; set; }
        public string lastUpdated { get; set; }
        public string mapQuest { get; set; }
        public string phone { get; set; }
        public string country { get; set; }
        public int parentId { get; set; }
        public string googleSearch { get; set; }
        public string linkedInSearch { get; set; }
        public string industry { get; set; }
        public string address { get; set; }
        public string AcctType { get; set; }
        public string websiteURL { get; set; }
        public string zipCode { get; set; }
    }
}
