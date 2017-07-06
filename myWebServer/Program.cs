//https://gist.github.com/aksakalli/9191056
//https://codehosting.net/blog/BlogEngine/post/Simple-C-Web-Server.aspx

using System;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Net.Http;
using System.Net.Sockets;

namespace myWebServer
{
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
                visible = value;
            }
            get
            {
                return visible;
            }
        }
        
        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        private static ConsoleHelper.SignalHandler signalHandler { get; set; }

        static string ServerGUID;
        static WebServer ws;        

        static void Main(string[] args)
        {
            signalHandler += HandleConsoleSignal;
            ConsoleHelper.SetSignalHandler(signalHandler, true);

            Visible = false;

            ServerGUID = args.Length > 0 ? args[0] + @"/" : "";
            string currentIP = IPTools.GetLocalIPAddress();
            string URI = "http://" + currentIP + ":8080/" + ServerGUID;
            ws = new WebServer(new string[] { URI }, SendResponse);
            ws.Run();

            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Altawin Utility Webserver");
            Console.WriteLine("Type ?help= after a command for detail information");

            bool isRun = true;

            while (isRun) {
                if (Visible)
                {
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.Out.WriteLine();
                    string request = Console.In.ReadLine();

                    switch (request.ToLower())
                    {
                        case "quit":
                        case "exit":
                            isRun = false;
                            break;
                        default:
                            HttpClient client = new HttpClient();                            
                            if (!request.Contains("/"))
                            {
                                request = "server/" + request + "?";
                            }
                            client.PostAsync(URI + request, null);
                            break;
                    }                    
                }
                else
                {
                    Thread.Sleep(1000);
                }
            }
            ws.Stop();
        }

        public static string SendResponse(HttpListenerRequest request, string GUID)
        {
            string Uri = request.RawUrl.Substring(1, request.RawUrl.IndexOf('?') - 1);
            Uri = Uri.Substring(ServerGUID.Length);
            string Part = Uri.Substring(0, Uri.IndexOf('/'));
            string Command = Uri.Substring(Uri.IndexOf('/')+1);
            CookieCollection Cookies = request.Cookies;
            string strCookies = string.Empty;
            if (request.Cookies != null)
            {
                foreach (Cookie item in request.Cookies)
                {
                    strCookies += item.Name + " = " + item.Value + "\r";
                }
            }            

            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();
            Console.Write("Request (" + DateTime.Now + ") from ");
            Console.WriteLine(request.UserHostName);
            Console.ResetColor();
            Console.WriteLine("Cookies");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(strCookies);
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
                    GetPricesFromExcel getPricesFromExcel;
                    if (GetPricesFromExcel.Sessions.ContainsKey(GUID))
                    {
                        getPricesFromExcel = GetPricesFromExcel.Sessions[GUID];
                    }
                    else
                    {
                        getPricesFromExcel = new GetPricesFromExcel(GUID);
                        GetPricesFromExcel.Sessions.Add(GUID, getPricesFromExcel);
                    }
                    Result = getPricesFromExcel.Execute(Command, request.QueryString);
                    
                    break;
                default:
                    Result = "Part [" + Part + "] is unknown";
                    break;
            }
            
            Console.WriteLine("Response (" + DateTime.Now + "):");
            Console.ForegroundColor = Result.StartsWith("Ok!") ? ConsoleColor.Green : ConsoleColor.Red;
            Console.WriteLine(Result);
            Console.ForegroundColor = ConsoleColor.Gray;

            return Result;
        }
        private static void HandleConsoleSignal(ConsoleSignal consoleSignal)
        {
            try
            {
                //GetPricesFromExcel.ExcelClose();
            }
            catch (Exception)
            {
            }
        }
    }
}
