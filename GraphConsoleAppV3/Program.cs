#region

using System;
using System.Collections.Generic;
using System.Net;

#endregion

namespace GraphConsoleAppV3
{
    using System.Data.Services.Client;
    using System.Web.Script.Serialization;

    public class Program
    {

        // Single-Threaded Apartment required for OAuth2 Authz Code flow (User Authn) to execute for this demo app
        [STAThread]
        private static void Main()
        {
            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            Console.WriteLine("Run operations for signed-in user, or in app-only mode.\n");
            Console.WriteLine("[a] - app-only\n[u] - as user\n[b] - both as user first, and then as app.\nPlease enter your choice:\n");

            ConsoleKeyInfo key = Console.ReadKey();
            switch (key.KeyChar)
            {
                case 'a':
                    Console.WriteLine("\nRunning app-only mode\n\n");
                    Requests.AppMode().Wait();
                    break;
                case 'b':
                    Console.WriteLine("\nRunning user mode, followed by app-only mode\n\n");
                    Requests.UserMode().Wait();
                    Requests.AppMode().Wait();
                    break;
                case 'u':
                    Console.WriteLine("\nRunning in user mode\n\n");
                    Requests.UserMode().Wait();
                    break;
                default:
                    WriteError("\nSelection not recognized. Running in user mode\n\n");
                    Requests.UserMode().Wait();
                    break;
            }

            //*********************************************************************************************
            // End of Demo Console App
            //*********************************************************************************************

            Console.WriteLine("\nCompleted at {0} \n Press Any Key to Exit.", DateTime.Now.ToUniversalTime());
            Console.ReadKey();
        }

        public static string ExtractErrorMessage(Exception exception)
        {
            List<string> errorMessages = new List<string>();


            string tabs = "\n";
            while (exception != null)
            {
                string requestIdLabel = "requestId";
                if (exception is DataServiceClientException &&
                    exception.Message.Contains(requestIdLabel))
                {
                    Dictionary<string, object> odataError =
                        new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(exception.Message);
                    odataError = (Dictionary<string, object>)odataError["odata.error"];
                    errorMessages.Insert(0, "\nRequest ID: " + odataError[requestIdLabel]);
                    errorMessages.Insert(1, "Date: " + odataError["date"]);
                }

                tabs += "    ";
                errorMessages.Add(tabs + exception.Message);
                exception = exception.InnerException;
            }

            return string.Join("\n", errorMessages);
        }

        public static void WriteError(string output, params object[] args)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine(output, args);
            Console.ResetColor();
        }
    }
}
