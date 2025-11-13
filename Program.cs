using System;
using System.Management;
using System.ServiceProcess;


namespace AutoUPS
{
    internal static class Program
    {
        static void Main()
        {
            ServiceBase.Run(new MyShutdownService());
        } 
    }
    }



