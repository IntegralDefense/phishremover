using System.ServiceProcess;
using Newtonsoft.Json;
using Office365;

namespace PhishRemover
{
    partial class PhishRemover : ServiceBase
    {
        WebServer webServer;

        public PhishRemover()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            base.OnStart(args);
            webServer = new WebServer();
            webServer.pages.Add("/delete", Delete);
            webServer.pages.Add("/restore", Restore);
            webServer.Start();
        }

        protected override void OnStop()
        {
            base.OnStop();
            webServer.Stop();
        }

        public static PageResult Delete(string json)
        {
            Email email = JsonConvert.DeserializeObject<Email>(json);
            ExchangeResult result = email.Delete();
            return new PageResult(200, result.message);
        }

        public static PageResult Restore(string json)
        {
            Email email = JsonConvert.DeserializeObject<Email>(json);
            ExchangeResult result = email.Restore();
            return new PageResult(200, result.message);
        }
    }
}
