using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace SharePointSearchTester
{
    /// <summary>
    /// Standalone console harness for testing the SharePoint 2016 Search REST call used by
    /// SharePointRestServiceObject.SearchByDocumentReference() in the K2 SmartObject broker,
    /// without needing the K2 SDK or a K2 environment.
    ///
    /// Usage:
    ///   SharePointSearchTester.exe --url "https://ssp16bbluehub.pbchbs.com/sites/EV/_api" --user "DOMAIN\svc-account" --pass "secret" --docref "202613500046"
    ///
    /// Arguments:
    ///   --url      SharePoint site _api base URL (no trailing slash)
    ///   --user     NTLM username (DOMAIN\user or user@domain)
    ///   --pass     NTLM password
    ///   --docref   Document Reference Number to search for (RefinableString00)
    /// </summary>
    internal class Program
    {
        private static readonly string[] SelectProperties = new[]
        {
            "RefinableString16", "RefinableString14", "RefinableString21", "RefinableString00", "RefinableString12",
            "RefinableString10", "RefinableDate00", "RefinableDate02", "RefinableString15", "RefinableDate01",
            "Title", "Path", "Description", "EditorOWSUSER", "LastModifiedTime", "CollapsingStatus", "DocId",
            "HitHighlightedSummary", "HitHighlightedProperties", "FileExtension", "ViewsLifeTime", "ParentLink",
            "FileType", "IsContainer", "SecondaryFileExtension", "DisplayAuthor", "PolicyGroup0OWSTEXT", "CertificateID0OWSTEXT"
        };

        private const int RowsPerPage = 50;

        private static int Main(string[] args)
        {
            string baseUrl, userName, password, docRef;

            try
            {
                var parsed = ParseArgs(args);
                baseUrl = parsed["url"].TrimEnd('/');
                userName = parsed["user"];
                password = parsed["pass"];
                docRef = parsed["docref"];
            }
            catch (Exception ex)
            {
                Console.WriteLine("Argument error: " + ex.Message);
                PrintUsage();
                return 1;
            }

            try
            {
                var results = SearchByDocumentReference(baseUrl, userName, password, docRef);

                Console.WriteLine("Matches found: " + results.Count);
                Console.WriteLine(new string('-', 60));

                int i = 1;
                foreach (var row in results)
                {
                    Console.WriteLine("Result #" + i++);
                    foreach (var kvp in row)
                    {
                        Console.WriteLine("  " + kvp.Key + " = " + kvp.Value);
                    }
                    Console.WriteLine(new string('-', 60));
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Search failed: " + ex.Message);
                if (ex.InnerException != null)
                    Console.WriteLine("Details: " + ex.InnerException.Message);
                return 2;
            }
        }

        private static Dictionary<string, string> ParseArgs(string[] args)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < args.Length - 1; i++)
            {
                if (args[i].StartsWith("--"))
                {
                    map[args[i].Substring(2)] = args[i + 1];
                    i++;
                }
            }

            string[] required = { "url", "user", "pass", "docref" };
            foreach (var key in required)
            {
                if (!map.ContainsKey(key))
                    throw new ArgumentException("Missing required argument: --" + key);
            }

            return map;
        }

        private static void PrintUsage()
        {
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  SharePointSearchTester.exe --url \"https://host/sites/site/_api\" --user \"DOMAIN\\svc-account\" --pass \"secret\" --docref \"202613500046\"");
        }

        /// <summary>
        /// Same logic as SharePointRestServiceObject.SearchByDocumentReference(), minus the K2 SDK dependency.
        /// Returns each matching document as a flat Key/Value dictionary of the selected properties.
        /// </summary>
        private static List<Dictionary<string, string>> SearchByDocumentReference(string baseUrl, string userName, string password, string documentReferenceNumber)
        {
            var results = new List<Dictionary<string, string>>();

            string kqlQuery = String.Format("RefinableString00:\"{0}\"", documentReferenceNumber);
            string querytext = Uri.EscapeDataString(kqlQuery);
            string selectProps = Uri.EscapeDataString(String.Join(",", SelectProperties));

            int startRow = 0;
            int totalRows = int.MaxValue;

            while (startRow < totalRows)
            {
                string url = String.Format(
                    "{0}/search/query?querytext='{1}'&rowlimit={2}&rowsperpage={2}&startrow={3}" +
                    "&trimduplicates=false&bypassresulttypes=true&processbestbets=false&processpersonalfavorites=false" +
                    "&clienttype='UI'&selectproperties='{4}'",
                    baseUrl, querytext, RowsPerPage, startRow, selectProps);

                Console.WriteLine("GET " + url);

                string json = ExecuteSearchRequest(url, userName, password);

                var serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                var response = serializer.Deserialize<Dictionary<string, object>>(json);

                var d = (Dictionary<string, object>)response["d"];
                var query = (Dictionary<string, object>)d["query"];
                var primary = (Dictionary<string, object>)query["PrimaryQueryResult"];
                var relevant = (Dictionary<string, object>)primary["RelevantResults"];

                totalRows = Convert.ToInt32(relevant["TotalRows"]);

                var table = (Dictionary<string, object>)relevant["Table"];
                var rowsWrapper = (Dictionary<string, object>)table["Rows"];
                var rows = (object[])rowsWrapper["results"];

                foreach (object rowObj in rows)
                {
                    var row = (Dictionary<string, object>)rowObj;
                    var cellsWrapper = (Dictionary<string, object>)row["Cells"];
                    var cells = (object[])cellsWrapper["results"];

                    var values = new Dictionary<string, string>();
                    foreach (object cellObj in cells)
                    {
                        var cell = (Dictionary<string, object>)cellObj;
                        string key = cell["Key"].ToString();
                        string value = cell["Value"] == null ? null : cell["Value"].ToString();
                        values[key] = value;
                    }

                    results.Add(values);
                }

                startRow += RowsPerPage;
            }

            return results;
        }

        private static string ExecuteSearchRequest(string url, string userName, string password)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.Accept = "application/json;odata=verbose";

            var credentials = new NetworkCredential(userName, password);
            var credentialCache = new CredentialCache();
            credentialCache.Add(request.RequestUri, "NTLM", credentials);
            request.Credentials = credentialCache;
            request.PreAuthenticate = true;
            request.UseDefaultCredentials = false;

            try
            {
                using (var httpResponse = (HttpWebResponse)request.GetResponse())
                using (var stream = httpResponse.GetResponseStream())
                using (var reader = new System.IO.StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                string details = wex.Message;
                if (wex.Response != null)
                {
                    using (var stream = wex.Response.GetResponseStream())
                    using (var reader = new System.IO.StreamReader(stream, Encoding.UTF8))
                    {
                        details = reader.ReadToEnd();
                    }
                }
                throw new ApplicationException("SharePoint search request failed: " + details, wex);
            }
        }
    }
}
