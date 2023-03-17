using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Web;
using DocumentFormat.OpenXml.Office2019.Excel.RichData2;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace ZephyrScaleTestCase
{
    class Program
    {
        public class Item
        {
            public string description { get; set; }
            public string expectedResult { get; set; }

            public string testData { get; set; }
        }
       
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            //Sandbox  string token = "Your Token";
            //PROD
            string token = "Your Token";
            string baseUrl = "https://api.zephyrscale.smartbear.com/v2";
            string projectKey = "PM";
            // Set up the HTTP client
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Set up Excel package
            var file = new FileInfo(@"C:\temp\PROD_ZephyrScaleTestCases.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets.Add("TestCases-" + projectKey);
                worksheet.Cells[1, 1].Value = "Zephyr Scale Test Case Key";
                worksheet.Cells[1, 2].Value = "Legacy Test Case ID";
                worksheet.Cells[1, 3].Value = "Zephyr Test Case ID";
                worksheet.Cells[1, 4].Value = "Initiative";
                worksheet.Cells[1, 5].Value = "ALM Creation Date";
                worksheet.Cells[1, 6].Value = "Test Case Type";
                worksheet.Cells[1, 7].Value = "Application";
                worksheet.Cells[1, 8].Value = "ALM Folder";
                worksheet.Cells[1, 9].Value = "Name";
                worksheet.Cells[1, 10].Value = "Description";
                worksheet.Cells[1, 11].Value = "Test Step URL";
                int row = 2;

                // Loop through all test cases and get the value of the custom field
                int offset = 0;
                const int pageSize = 10;
                int total = pageSize;
                while (offset < total)
                {
                    try
                    {
                        // Get the list of test cases
                        var testCasesResponse = await httpClient.GetAsync($"{baseUrl}/testcases?projectKey={projectKey}&startAt={offset}&maxRecords={pageSize}");
                        var testCasesJson = await testCasesResponse.Content.ReadAsStringAsync();


                        // 

                        testCasesJson = testCasesJson.Replace("\t", "").Replace("\r\n", "");
                        var loop = true;
                        do
                        {
                            try
                            {
                                var m = JsonConvert.DeserializeObject<T>(testCasesJson);
                                loop = false;
                            }
                            catch (JsonReaderException ex)
                            {
                                var position = ex.LinePosition;
                                var invalidChar = testCasesJson.Substring(position - 2, 2);
                                invalidChar = invalidChar.Replace("\"", "'");
                                testCasesJson = $"{testCasesJson.Substring(0, position - 1)}{invalidChar}{testCasesJson.Substring(position)}";
                            }
                        } while (loop);

                        var testCases = JsonConvert.DeserializeObject<JObject>(testCasesJson);
                        total = (int)testCases["total"];

                        ///////////////////////////
                     //   total = 10; //Prem ToDebug
                        //////////////////////
                        
                        // Loop through the test cases and get the custom field value
                        var testCasesArray = (JArray)testCases["values"];
                        foreach (var testCase in testCasesArray)
                        {
                            string testCaseId = testCase["id"].ToString();
                            string testCaseKey = testCase["key"].ToString();
                            string objective = testCase["objective"].ToString();
                            string name = testCase["name"].ToString();
                            // Get the details of the test case
                            //  var testCaseResponse = await httpClient.GetAsync($"{baseUrl}/testcases/{testCaseId}");
                            //   var testCaseJson = await testCaseResponse.Content.ReadAsStringAsync();
                            var testCaseDetails = JsonConvert.DeserializeObject<JObject>(testCase.ToString());

                            // Get the value of the custom field "Legacy Test Case ID"
                            var customFields = (JObject)testCaseDetails["customFields"];
                            string legacyTestId = customFields["Legacy Test Case ID"]?.ToString();
                            string initativeId = customFields["Initiative"]?.ToString();
                            string creationDate = customFields["Creation Date"]?.ToString();
                            string testType = customFields["Test Type"]?.ToString();
                            string application = customFields["Application"]?.ToString();
                            string almFolderPath = customFields["ALM Folder Path"]?.ToString();

                            var testStepDetails = (JObject)testCaseDetails["testScript"];
                            string testStepUrl = testStepDetails["self"]?.ToString();

                            var testStepsResponse = await httpClient.GetAsync(testStepUrl);
                            var testStepsJson = await testStepsResponse.Content.ReadAsStringAsync();
                            var testSteps = JsonConvert.DeserializeObject<JObject>(testStepsJson);

                            Console.WriteLine("Zephyr Testcase Key:" + testCaseKey);
                            Console.WriteLine("ALM Testcase:" + legacyTestId);
                            // Write the values to the Excel worksheet
                            worksheet.Cells[row, 1].Value = testCaseId;
                            worksheet.Cells[row, 2].Value = testCaseKey;
                            worksheet.Cells[row, 3].Value = legacyTestId;
                            worksheet.Cells[row, 4].Value = initativeId;
                            worksheet.Cells[row, 5].Value = creationDate;
                            worksheet.Cells[row, 6].Value = testType;
                            worksheet.Cells[row, 7].Value = application;
                            worksheet.Cells[row, 8].Value = almFolderPath;
                            worksheet.Cells[row, 9].Value = name;
                            worksheet.Cells[row, 10].Value = objective;
                            worksheet.Cells[row, 11].Value = testStepUrl;

                            var testStepsArray = (JArray)testSteps["values"];
                            //  var testStepsArrayInline = JsonConvert.DeserializeObject<JObject>(testStepsArray.ToString());

                            // List<object> list = testStepsArray.ToList<object>();
                            int stepsCount;
                            try
                            {
                                stepsCount = int.Parse(testSteps["total"].ToString());
                            }
                            catch (Exception ex)
                            {
                                stepsCount = 0;
                            }
                            int j = 0;
                            int k = 12;
                            for (int i=12; i<(stepsCount+12);i++)
                            {
                                try
                                {
                                    var testStepsArrayInlineFields = (JObject)testStepsArray[j]["inline"];

                                    //  var testStepsInlineArray = (JArray)testStepsArrayInline[0]["description"];
                                    worksheet.Cells[1, k].Value = "Step Number -" + j;
                                    worksheet.Cells[1, k + 1].Value = "Description - " + j;
                                    worksheet.Cells[1, k + 2].Value = "Expected Result - " + j;

                                    worksheet.Cells[row, k].Value =   testStepsArrayInlineFields["description"]?.ToString().Replace("\"", "");
                                    worksheet.Cells[row, k + 1].Value = testStepsArrayInlineFields["testData"]?.ToString().Replace("\"", "");
                                    worksheet.Cells[row, k + 2].Value = testStepsArrayInlineFields["expectedResult"]?.ToString().Replace("\"", "");
                                    j = j + 1;
                                    k = k + 3;
                                }
                                catch (Exception ex)
                                {
                                    worksheet.Cells[row, k].Value = "Error";
                                    worksheet.Cells[row, k + 1].Value = "Error";
                                    worksheet.Cells[row, k + 2].Value = "Error";
                                }
                            }


                            row++;
                            package.Save();
                        }

                        offset += pageSize;
                    }
                    catch (Exception ex)
                    {
                        int milliseconds = 2000;
                        Thread.Sleep(milliseconds);
                    }
                }

                // Save the Excel package
                package.Save();
            }
        }
    }
}
