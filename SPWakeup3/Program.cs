using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using System.Web;
using System.Collections;

namespace SPWakeup3
{
    public class InitValues
    {
        public InitValues()
        { }

        public void AppendExclude(string newAddress)
        {
            if (!excludeArray.Contains(newAddress.ToLower()))
            {
                excludeArray.Add(newAddress.ToLower());
            }
        }
        //public string logFile = "";
        //public string excludeFile = "";
        public ArrayList excludeArray = new ArrayList();
        public string userName = "";
        public string password = "";
        public string domain = "";
        public string authType = "NTLM";
        public string email = "";
        public string mailSite = "";
        public bool verbose = false;
        public bool run = true;
        //public string mailServer = "";
    }

    public class SiteCollection
    {
        public SiteCollection(string newSite, ArrayList newIgnoreArray, bool verbose)
        {
            ignoreArray = newIgnoreArray;
            try
            {
                using (SPSite currentSite = new SPSite(newSite))
                {
                    if (!ignoreArray.Contains(currentSite.RootWeb.Url))
                    {
                        if (!subWebs.Contains(currentSite.RootWeb.Url))
                        {
                            subWebs.Add(currentSite.RootWeb.Url);
                            URL = currentSite.RootWeb.Url;
                            Console.WriteLine("Added Site Collection: " + currentSite.RootWeb.Url);
                        }
                        FindAllSubWebs(currentSite.RootWeb, verbose);
                        switch (webCount)
                        {
                            case 0:
                                break;
                            case 1:
                                Console.WriteLine("Found 1 sub-web.");
                                break;
                            default:
                                Console.WriteLine("Found " + webCount.ToString() + " sub-webs.");
                                break;
                        }
                    }
                }
            }
            catch
            { Console.WriteLine("Error opening Site Collection: " + newSite); }
        }

        private void FindAllSubWebs(SPWeb checkWeb, bool verbose)
        {
            try
            {
                foreach (SPWeb newWeb in checkWeb.Webs)
                {
                    if (!ignoreArray.Contains(newWeb.Url))
                    {
                        if (!subWebs.Contains(newWeb.Url))
                        {
                            subWebs.Add(newWeb.Url);
                            webCount++;
                            if (verbose)
                            {
                                Console.WriteLine("Added Sub Web: " + newWeb.Url);
                            }
                        }
                        FindAllSubWebs(newWeb, verbose);
                    }
                }
            }
            catch
            { Console.WriteLine("Error getting subwebs."); }
        }

        public void AppendWeb(SPWeb newWeb)
        {
        if ((!subWebs.Contains(newWeb.Url)) && (!ignoreArray.Contains(newWeb.Url)))
        {
        
        }
        }

        ArrayList ignoreArray = new ArrayList();
        public ArrayList subWebs = new ArrayList();
        public int webCount = 0;
        public string URL = "";
    }

    public class Log
    {
        public Log()
        { }
                
        public void AppendEntry(string newEntry)
        {
            logArray.Add(newEntry);
            Console.WriteLine(newEntry);
        }

        public void AddSite(string newURL)
        {
            siteCount++;
        }
        
        public void AddSuccess()
        {
            successCount++;
        }

        public void AddFailure(string failedURL, string errorMessage)
        {
            failCount++;
            string[] failure = {failedURL, errorMessage};
            failArray.Add(failure);
            Console.WriteLine("Failed waking site: " + failedURL);
            Console.WriteLine("Error: " + errorMessage);
        }

        public ArrayList logArray = new ArrayList();
        public ArrayList failArray = new ArrayList();
        int siteCount = 0;
        public int successCount = 0;
        public int failCount = 0;
    }

    class Program
    {
        public static InitValues initVals = new InitValues();
        public static Log log = new Log();
        public static ArrayList ignoreArray = new ArrayList();
        
        static void Main(string[] args)
        {
            log.AppendEntry("Started running at " + DateTime.Now.ToString());
            log.AppendEntry("Running on the system " + System.Environment.MachineName);

            GetInitValues(args);

            if (initVals.run)
            {
                ArrayList siteCollArray = new ArrayList();
                try
                {
                    SPFarm farm = SPFarm.Local;
                    SPWebService service = farm.Services.GetValue<SPWebService>("");
                    foreach (SPWebApplication webApp in service.WebApplications)
                    {
                        try
                        {
                            foreach (SPSite currentsite in webApp.Sites)
                            {
                                if (NotOnIgnoreList(currentsite.Url))
                                {
                                    SiteCollection newSite = new SiteCollection(currentsite.Url, ignoreArray, initVals.verbose);
                                    siteCollArray.Add(newSite);
                                    log.AddSite(currentsite.Url);

                                    //Set our outgoing email site to the first site collection found
                                    if (initVals.mailSite == "")
                                    { initVals.mailSite = currentsite.Url; }
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch
                { Console.WriteLine("Trouble getting handle on Sharepoint Farm. Must be run on a Sharepoint/WSS server."); }

                foreach (SiteCollection currentSite in siteCollArray)
                {
                    ArrayList subWebs = currentSite.subWebs;
                    if (subWebs.Count > 0)
                    {
                        if (currentSite.webCount == 0)
                        {
                            log.AppendEntry("");
                            log.AppendEntry("Waking site under the address " + currentSite.URL);
                        }
                        else
                        {
                            log.AppendEntry("");
                            log.AppendEntry("Waking " + (currentSite.webCount + 1).ToString() + " sites under the address " + currentSite.URL);
                        }

                        int x = 0;
                        foreach (string URL in subWebs)
                        {
                            x++;
                            bool success = GetWebPage(URL, initVals.userName, initVals.password, initVals.domain, initVals.authType);
                            if (success)
                            {
                                if (initVals.verbose)
                                {
                                    log.AppendEntry("Woke: " + URL);
                                }
                                else
                                {
                                    drawTextProgressBar(x, subWebs.Count);
                                    //Console.Write(".");
                                }
                            }
                        }
                        
                        Console.WriteLine("");
                        Console.WriteLine("");
                    }
                }

                if (initVals.email != "")
                {
                    string br = "<br>\n";

                    string toAddress = initVals.email;

                    string subject = "SPWakeup finished running with " + log.successCount.ToString() + " successes and " + log.failCount.ToString() + " failures on " +  System.Environment.MachineName;

                    string body = "SPWakeup finished running at " + DateTime.Now.ToString() + " on " + System.Environment.MachineName + br + br;
                    switch (log.successCount)
                    {
                        case 0:
                            body += "Did not wake any sites." + br;
                            break;
                        case 1:
                            body += "Sucessfully woke 1 site1." + br;
                            break;
                        default:
                            body += "Successfully woke " + log.successCount.ToString() + " sites." + br;
                            break;
                    }
                    switch (log.failCount)
                    {
                        case 0:
                            body += "No failures." + br;
                            break;
                        case 1:
                            body += "Failed to wake 1 site." + br;
                            break;
                        default:
                            body += "Failed to wake " + log.failCount.ToString() + " sites." + br;
                            break;
                    }
                    if (log.failCount > 0)
                    {
                        body += br + "Failed Site List:" + br;
                        foreach (string[] currentFail in log.failArray)
                        {
                            body += currentFail[0].ToString() + br;
                            body += "Failed with error: " + currentFail[1].ToString() + br;
                        }
                    }

                    if (initVals.verbose)
                    {
                        body += br + br + "Run time log below: " + br + br;

                        foreach (string entry in log.logArray)
                        { body += entry + br; }
                    }

                    EmailResults(toAddress, subject, body);
                }
            }
            else
            {
                ArrayList help = new ArrayList();
                help.Add("If SPWakeup is run without any run time options, it will attempt to find and");
                help.Add("open every site on the local server's Sharepoint farm.");
                help.Add("You can vary this behaviour by using the run time options listed below.");
                help.Add("These options are not case sensitive.");
                help.Add("");
                help.Add("Available run-time options are: ");
                help.Add("-Exclude: Excludes the listed Site Collection URL from being woken.");
                help.Add("Can be used more than once. Example:");
                help.Add("spwakeup3.exe -Exclude:http://badsite.com -Exclude:http://badsite2.com");
                help.Add("-Email: An email address that should be sent a log of the results.");
                help.Add("-UserName: Name of the account that should be used to browse the sites.");
                help.Add("If no user name is set, sites are accessed under the current account.");
                help.Add("-Domain: The domain of the account specified above.");
                help.Add("-Password: The password that should be used to browse the sites.");
                help.Add("You only need to set this option if you are also specifying an account.");
                help.Add("-Authentication: The type of authentication used to browse the sites.");
                help.Add("By default NTLM authentication is used.");
                help.Add("-Verbose If this flag is set, every site is listed by URL.");
                help.Add("By default only the Total number of sites is listed.");
                help.Add("[Warning] If you use the Email option in conjunction with Verbose mode, the");
                help.Add("resulting email may be cut-off after 2,048 characters.  This is due to a bug");
                help.Add("with Sharepoint's SPUtility.SendEmail function.");
                help.Add("-Help Shows this help screen.");

                foreach (string newLine in help)
                { log.AppendEntry(newLine); }
            }
        }

        private static bool NotOnIgnoreList(string URL)
        {
            bool notOnIgnoreList = true;

            string checkURL = URL.ToLower();
            foreach (string currentExcludeURL in initVals.excludeArray)
            { 
                if (checkURL.StartsWith(currentExcludeURL))
                {
                    notOnIgnoreList = false;
                }
            }

            return notOnIgnoreList;
        }

        private static void drawTextProgressBar(int progress, int total)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 32;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 30.0f / total;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw unfilled part
            for (int i = position; i <= 31; i++)
            {
                Console.BackgroundColor = ConsoleColor.Black;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            Console.CursorLeft = 35;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write(progress.ToString() + " of " + total.ToString() + "    "); //blanks at the end remove any excess
        }


        private static void EmailResults(string toAddress, string subject, string body)
        {
            try
            {
                if (initVals.mailSite != "")
                {
                    using (SPSite rootSite = new SPSite(initVals.mailSite))
                    {
                        Console.WriteLine("Emailing results to " + toAddress);
                        SPUtility.SendEmail(rootSite.RootWeb, false, false, toAddress, subject, body);
                    }
                }
                else
                { Console.WriteLine("Error sending email, could not find working site."); }
            }
            catch (Exception ex)
            { 
                Console.WriteLine("Error emailing results to " + toAddress);
                Console.WriteLine("Error Message: " + ex.Message.ToString());
            }
        }

        private static void GetInitValues(string[] args)
        {
            foreach (string currentArg in args)
            {
                string checkVal = currentArg.ToLower();
                if (!checkVal.Contains(":"))
                { checkVal = checkVal + ":"; }

                switch (checkVal.Substring(0, checkVal.IndexOf(":")))
                {
                    case "-email":
                        initVals.email = currentArg.Substring(7);
                        break; 
                    //case "-mailserver":
                        //initVals.mailServer = currentArg.Substring(12);
                        //break;
                    //case "-log":
                    //    initVals.logFile = currentArg.Substring(5);
                    //    break;
                    //case "-excludefile":
                    //    initVals.excludeFile = currentArg.Substring(8);
                    //    break;
                    case "-exclude":
                        initVals.AppendExclude(currentArg.Substring(9));
                        log.AppendEntry("Excluding the site: " + currentArg.Substring(9));
                        break;
                    case "-username":
                        initVals.userName = currentArg.Substring(10);
                        break;
                    case "-domain":
                        initVals.domain = currentArg.Substring(8);
                        break;
                    case "-password":
                        initVals.password = currentArg.Substring(10);
                        break;
                    case "-authentication":
                        initVals.authType = currentArg.Substring(16);
                        break;
                    case "-verbose":
                        initVals.verbose = true;
                        log.AppendEntry("Setting verbose mode.");
                        break;
                    case "-help":
                        initVals.run = false;
                        log.AppendEntry("Displaying Help");
                        log.AppendEntry("");
                        break;
                    default:
                        initVals.run = false;
                        StringBuilder errorMessage = new StringBuilder();
                        log.AppendEntry("Invalid Run time option entered: " + checkVal);
                        log.AppendEntry("");
                        break;
                }
            }
        }

        public static bool GetWebPage(string URL, string UserName, string Password, string Domain, string AuthType)
        {
            try
            {
                // Create a request for the URL.  
                WebRequest request = WebRequest.Create(URL);

                request.Proxy = null;

                // If required by the server, set the credentials.
                if ((UserName == null) || (UserName == ""))
                {
                    request.Credentials = CredentialCache.DefaultCredentials;
                }
                else
                {
                    CredentialCache myCache = new CredentialCache();
                    myCache.Add(new Uri(URL), AuthType, new NetworkCredential(UserName, Password, Domain));
                    request.Credentials = myCache;
                }

                // Get the response.
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                // Display the status.
                // Console.WriteLine (response.StatusDescription);

                // Get the stream containing content returned by the server.
                Stream dataStream = response.GetResponseStream();

                // Open the stream using a StreamReader for easy access.
                StreamReader reader = new StreamReader(dataStream);

                string responseFromServer = reader.ReadToEnd();

                // Cleanup the streams and the response.
                reader.Close();
                dataStream.Close();
                response.Close();
                log.AddSuccess();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                log.AddFailure(URL, ex.Message.ToString());
                return false;
            }
        }
    }
}
