//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft_Graph_ASPNET_Excel_Donations.Models;
using Newtonsoft.Json;
using System.Drawing;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace Microsoft_Graph_ASPNET_Excel_Donations
{
    public class ExcelApiHelper
    {
        private static string restURLBase = "https://graph.microsoft.com/beta/me/drive/items/";
        private static string fileId = null;

        private static async Task UpdateSummaryTable(string tableName, string field1Value, string field2Value, string worksheetsEndpoint, HttpClient client)
        {
            var summaryTableRowJson = "{" +
                    "'values': [['" + field1Value + "', '" + field2Value + "']]" +
                "}";
            var summaryTableRowContentPostBody = new StringContent(summaryTableRowJson, System.Text.Encoding.UTF8);
            summaryTableRowContentPostBody.Headers.Clear();
            summaryTableRowContentPostBody.Headers.Add("Content-Type", "application/json");
            var summaryTableRowResponseMessage = await client.PostAsync(worksheetsEndpoint + "tables('" + tableName + "')/rows", summaryTableRowContentPostBody);
        }


        public static async Task<List<Donation>> GetDonations(string accessToken)
        {
            List<Donation> donations = new List<Donation>();

            using (var client = new HttpClient())
            {
                //client.BaseAddress = new Uri(restURLBase);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                //Find file id
                var serviceEndpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children";
                var filesResponse = await client.GetAsync(serviceEndpoint + "?$select=name,id");

                var filesContent = await filesResponse.Content.ReadAsStringAsync();

                JObject parsedResult = JObject.Parse(filesContent);

                foreach (JObject file in parsedResult["value"])
                {

                    var name = (string)file["name"];
                    if (name.Contains("WoodGroveBankExpenseTrendsWorkbook.xlsx"))
                    {
                        fileId = (string)file["id"];
                        restURLBase = "https://graph.microsoft.com/beta/me/drive/items/" + fileId + "/workbook/worksheets('Donations')/";

                    }
                }         

                HttpResponseMessage response = await client.GetAsync(restURLBase + "tables('DonationsTable')/Rows");
                if (response.IsSuccessStatusCode)
                {
                    string resultString = await response.Content.ReadAsStringAsync();

                    dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(resultString);
                    JArray y = x.value;

                    donations = BuildList(donations, y);
                }
            }

            return donations;
        }

        private static List<Donation> BuildList(List<Donation> donations, JArray y)
        {
            foreach (var item in y.Children())
            {
                var itemProperties = item.Children<JProperty>();

                //Get element that holds row collection
                var element = itemProperties.FirstOrDefault(xx => xx.Name == "values");
                JProperty index = itemProperties.FirstOrDefault(xxx => xxx.Name == "index");

                //The string array of row values
                JToken values = element.Value;

                //linq query to get rows from results
                var stringValues = from stringValue in values select stringValue;
                //rows
                foreach (JToken thing in stringValues)
                {
                    IEnumerable<string> rowValues = thing.Values<string>();

                    //Cast row value collection to string array
                    string[] stringArray = rowValues.Cast<string>().ToArray();

                    try
                    {
                        Donation donation = new Donation(
                             DateTime.FromOADate(Convert.ToDouble(stringArray[0])),
                             Convert.ToDecimal(stringArray[1]),
                             stringArray[2],
                             stringArray[3]);
                        donations.Add(donation);
                    }
                    catch (FormatException f)
                    {
                        //Handle exception
                    }
                }
            }

            return donations;

        }

        public static async Task<Donation> CreateDonation(
                                                 string accessToken,
                                                 string donationDate,
                                                 string donationAmount,
                                                 string organization)
        {
            Donation newDonation = new Donation();


            using (var client = new HttpClient())
            {


                client.BaseAddress = new Uri(restURLBase);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                //Find file id
                var serviceEndpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children";
                var filesResponse = await client.GetAsync(serviceEndpoint + "?$select=name,id");

                var filesContent = await filesResponse.Content.ReadAsStringAsync();

                JObject parsedResult = JObject.Parse(filesContent);

                foreach (JObject file in parsedResult["value"])
                {

                    var name = (string)file["name"];
                    if (name.Contains("WoodGroveBankExpenseTrendsWorkbook.xlsx"))
                    {
                        fileId = (string)file["id"];
                        restURLBase = "https://graph.microsoft.com/beta/me/drive/items/" + fileId + "/workbook/worksheets('Donations')/";

                    }
                }


                using (var request = new HttpRequestMessage(HttpMethod.Post, restURLBase))
                {

                    var donationMonth = "=TEXT([DATE], \"mmm - yyyy\")";
                    
                    //Create 2 dimensional array to hold the row values to be serialized into json
                    object[,] valuesArray = new object[1, 4] { { donationDate, donationAmount, organization, donationMonth } };

                    //Create a container for the request body to be serialized
                    RequestBodyHelper requestBodyHelper = new RequestBodyHelper();
                    requestBodyHelper.index = null;
                    requestBodyHelper.values = valuesArray;

                    //Serialize the final request body
                    string postPayload = JsonConvert.SerializeObject(requestBodyHelper);

                    //Add the json payload to the POST request
                    request.Content = new StringContent(postPayload, System.Text.Encoding.UTF8);



                    using (HttpResponseMessage response = await client.PostAsync("tables('DonationsTable')/rows", request.Content))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string resultString = await response.Content.ReadAsStringAsync();
                            dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(resultString);
                        }

                    }
                    DateTime donationDateTime = Convert.ToDateTime(donationDate);
                    var donationMonthString = donationDateTime.ToString("MMM - yyyy");

                    await UpdateSummaryTable("DonationsByOrgTable", organization, donationAmount, restURLBase,  client);
                    await UpdateSummaryTable("DonationsByMonthTable", donationMonthString, donationAmount, restURLBase,  client);
                }
            }
            return newDonation;
        }

    }

    public class RequestBodyHelper
    {
        public object index;
        public object[,] values;
    }
}