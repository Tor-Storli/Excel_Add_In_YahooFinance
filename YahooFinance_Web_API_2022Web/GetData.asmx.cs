
using System;
using System.Collections.Generic;
using System.Web.Services;
using System.Net;
using System.IO;

namespace YahooFinance_Web_API_2022Web
{
    /// <summary>
    /// Summary description for GetData
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]

    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    [System.Web.Script.Services.ScriptService]
    public class GetData: WebService
    {
       // private static readonly HttpClient client = new HttpClient();

       [WebMethod]
        public List<string> GetDataFromWebApi(string ticker, string range)
        {

            //Console.WriteLine(ticker);
           // Console.WriteLine(range);

            string baseUrl = "https://query1.finance.yahoo.com/v8/finance/chart/";
            string queryUrl = String.Format("{0}?region=US&lang=en-US&includePrePost=false&interval=1d&range={1}&corsDomain=finance.yahoo.com&.tsrc=finance", ticker, range);
           
            string url = baseUrl + queryUrl;

            // Create a request for the URL.        
            WebRequest request = WebRequest.Create(url);
            
            // If required by the server, set the credentials.
            request.Credentials = CredentialCache.DefaultCredentials;
            
            // Get the response.
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            
            // Display the status.
            Console.WriteLine(response.StatusDescription);
           
            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
           
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            
            // Display the content.
            Console.WriteLine(responseFromServer);

            // Cleanup the streams and the response.
            reader.Close();
            dataStream.Close();
            response.Close();

            List<string> lstData = new List<string>
            {
                responseFromServer
            };

            return lstData;

        }
       
    }
}
