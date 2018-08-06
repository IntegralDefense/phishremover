using System;
using Newtonsoft.Json;
using Office365;

namespace PhishRemoverConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.Write("Action (quit|delete|restore): ");
                string action = Console.ReadLine().ToLower();

                if (action.StartsWith("q"))
                {
                    break;
                }

                Console.Write("Recipient: ");
                string recipient = Console.ReadLine();
                recipient = recipient.Replace("\\", "\\\\");
                recipient = recipient.Replace("\"", "\\\"");

                Console.Write("MessageId: ");
                string message_id = Console.ReadLine();
                message_id = message_id.Replace("\\", "\\\\");
                message_id = message_id.Replace("\"", "\\\"");

                string json = "{\"recipient\":\"" + recipient + "\",\"message_id\":\"" + message_id + "\"}";

                Email email = JsonConvert.DeserializeObject<Email>(json);

                ExchangeResult result = null;
                if (action.StartsWith("d"))
                {
                    result = email.Delete();
                } else if (action.StartsWith("r"))
                {
                    result = email.Restore();
                } else
                {
                    Console.WriteLine("invalid action");
                    continue;
                }
                Console.WriteLine(result.message);
            }
        }
    }
}
