using System.Runtime.InteropServices;

namespace myWebServer
{
    internal static class ConsoleHelper
    {
        [DllImport("Kernel32", EntryPoint = "SetConsoleCtrlHandler")]
        public static extern bool SetSignalHandler(SignalHandler handler, bool add);
        internal delegate void SignalHandler(ConsoleSignal consoleSignal);
    }
}
