//https://gist.github.com/aksakalli/9191056
//https://codehosting.net/blog/BlogEngine/post/Simple-C-Web-Server.aspx

using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Net;
//using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace myWebServer
{
    internal delegate void SignalHandler(ConsoleSignal consoleSignal);

    internal enum ConsoleSignal
    {
        CtrlC = 0,
        CtrlBreak = 1,
        Close = 2,
        LogOff = 5,
        Shutdown = 6
    }

    internal static class ConsoleHelper
    {
        [DllImport("Kernel32", EntryPoint = "SetConsoleCtrlHandler")]
        public static extern bool SetSignalHandler(SignalHandler handler, bool add);
    }

    class Program
    {
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private static IntPtr programHandle = GetConsoleWindow();
        private static bool visible;
        public static bool Visible {
            set
            {
                ShowWindow(programHandle, value ? SW_SHOW : SW_HIDE);
                visible = true;
            }
            get
            {
                return visible;
            }
        }
        
        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        private static SignalHandler signalHandler;

        static string GUID;
        static WebServer ws;
        static void Main(string[] args)
        {
            signalHandler += HandleConsoleSignal;
            ConsoleHelper.SetSignalHandler(signalHandler, true);

            Visible = false;

            GUID = args.Length > 0 ? args[0] + @"/" : "";
            ws = new WebServer(SendResponse, "http://localhost:8080/" + GUID);
            ws.Run();

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Altawin Utility Webserver");
            Console.WriteLine("Type ?help= after a command for detail information");

            while (true) {
                if (Visible)
                {
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.Out.WriteLine();
                    string request = Console.In.ReadLine();
                    
                    if (request.ToLower().Equals("quit") || request.ToLower().Equals("exit"))
                    {
                        break;
                    }
                    else
                    {
                        //HttpClient client = new HttpClient();
                        //client.UploadString("http://localhost:8080/" + GUID, request);
                    }
                }
                else
                {
                    Thread.Sleep(1000);
                }
            }
            ws.Stop();
        }

        private static string sendResponse(string Part, string Command, string UserHostName, System.Collections.Specialized.NameValueCollection QueryString)
        {
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();
            Console.Write("Request (" + DateTime.Now + ") from ");
            Console.WriteLine(UserHostName);
            Console.ResetColor();
            Console.WriteLine("Part:  ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(Part);
            Console.ResetColor();
            Console.WriteLine("Command:  ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(Command);
            Console.ResetColor();
            
            Console.WriteLine("Параметри:");
            Console.ForegroundColor = ConsoleColor.Yellow;
            foreach (string paramName in QueryString.AllKeys)
            {
                object[] paramValues = QueryString.GetValues(paramName);
                for (int i = 0; i < paramValues.Length; i++)
                {
                    Console.WriteLine(paramName + "=" + QueryString.GetValues(paramName).GetValue(i).ToString());
                }
            }
            Console.ResetColor();
            
            string Result = "";

            switch (Part.ToLower())
            {
                case "server":
                    Result = ws.Execute(Command, QueryString);
                    break;
                case "getpricesfromexcel":
                    Result = getPricesFromExcel.Execute(Command, QueryString);
                    break;
                default:
                    Result = "Part [" + Part + "] is unknown";
                    break;
            }

            Console.WriteLine("Response (" + DateTime.Now + "):");
            Console.ForegroundColor = Result.StartsWith("Ok!") ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine(Result);
            return Result;
        }

        public static string SendResponse(HttpListenerRequest request)
        {
            string Uri = request.RawUrl.Substring(1, request.RawUrl.IndexOf('?') - 1);
            Uri = Uri.Substring(GUID.Length);
            string Part = Uri.Substring(0, Uri.IndexOf('/'));
            string Command = Uri.Substring(Uri.IndexOf('/')+1);

            //return sendResponse(Part, Command, request.UserHostName, request.QueryString);
            
            #region Old version

            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();
            Console.Write("Request (" + DateTime.Now + ") from ");
            Console.WriteLine(request.UserHostName);
            Console.ResetColor();
            Console.WriteLine("Part:  ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(Part);
            Console.ResetColor();
            Console.WriteLine("Command:  ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(Command);
            Console.ResetColor();

            Console.WriteLine("Параметри:");
            Console.ForegroundColor = ConsoleColor.Yellow;
            foreach (string paramName in request.QueryString.AllKeys)
            {
                object[] paramValues = request.QueryString.GetValues(paramName);
                for (int i = 0; i < paramValues.Length; i++)
                {
                    Console.WriteLine(paramName + "=" + request.QueryString.GetValues(paramName).GetValue(i).ToString());
                }
            }
            Console.ResetColor();
            
            string Result = "";

            switch (Part.ToLower())
            {
                case "server":
                    Result = ws.Execute(Command, request.QueryString);
                    break;
                case "getpricesfromexcel":
                    Result = getPricesFromExcel.Execute(Command, request.QueryString);
                    
                    break;
                default:
                    Result = "Part [" + Part + "] is unknown";
                    break;
            }
            
            Console.WriteLine("Response (" + DateTime.Now + "):");
            Console.ForegroundColor = Result.StartsWith("Ok!") ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine(Result);
            return Result;
            
            #endregion

        }
        private static void HandleConsoleSignal(ConsoleSignal consoleSignal)
        {
            try
            {
                getPricesFromExcel.ExcelClose();
            }
            catch (Exception)
            {
            }
        }
    }
}
