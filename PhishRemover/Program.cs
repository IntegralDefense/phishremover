using System.ServiceProcess;

namespace PhishRemover
{
    class Program
    {
        static void Main(string[] args)
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] { new PhishRemover() };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
