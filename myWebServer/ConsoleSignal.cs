using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace myWebServer
{
    internal enum ConsoleSignal
    {
        CtrlC = 0,
        CtrlBreak = 1,
        Close = 2,
        LogOff = 5,
        Shutdown = 6
    }
}
