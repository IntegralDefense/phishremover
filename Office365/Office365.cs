using System.Collections.Generic;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;
using System;
using Newtonsoft.Json;

namespace Office365
{
    public enum StatusCode { Success = 200, RecipientNotFound = 210, MessageNotFound = 211, Error = 500 };

    public class ExchangeResult
    {
        class ExchangeResultLog
        {
            public string address;
            public StatusCode code;
            public string message;

            public ExchangeResultLog(string address, StatusCode code, string message)
            {
                this.address = address;
                this.code = code;
                this.message = message;
            }
        }
        List<ExchangeResultLog> logs = new List<ExchangeResultLog>();
        public string message { get { return JsonConvert.SerializeObject(logs); } }

        public ExchangeResult() { }

        public ExchangeResult(string address, StatusCode code, string message)
        {
            Log(address, code, message);
        }

        public void Log(string address, StatusCode code, string message)
        {
            logs.Add(new ExchangeResultLog(address, code, message));
        }

        public void Log(ExchangeResult exchangeResults)
        {
            if (exchangeResults != null)
            {
                logs.AddRange(exchangeResults.logs);
            }
        }
    }

    public class Email
    {
        public string recipient = "";

        private string _message_id = "";
        public string message_id
        {
            get { return _message_id; }
            set
            {
                _message_id = value;
                if (!_message_id.StartsWith("<")) _message_id = "<" + _message_id;
                if (!_message_id.EndsWith(">")) _message_id += ">";
            }
        }

        public ExchangeResult Delete()
        {
            ExchangeResult result = new ExchangeResult();

            // recipient can contain multiple addresses seperated by semi colons
            foreach (string address in recipient.Split(';'))
            {
                result.Log(ExecuteOperation(EmailOperation.Delete, address));
            }

            return result;
        }

        public ExchangeResult Restore()
        {
            ExchangeResult result = new ExchangeResult();

            // recipient can contain multiple addresses seperated by semi colons
            foreach (string address in recipient.Split(';'))
            {
                result.Log(ExecuteOperation(EmailOperation.Restore, address));
            }

            return result;
        }

        static List<WebCredentials> credentials
        {
            get
            {
                List<WebCredentials> creds = new List<WebCredentials>();
                string[] accounts = ConfigurationManager.AppSettings["accounts"].Split(',');
                foreach (string account in accounts)
                {
                    string[] part = account.Split(':');
                    creds.Add(new WebCredentials(part[0], part[1]));
                }
                return creds;
            }
        }

        enum EmailOperation { Delete, Restore }

        ExchangeResult ExecuteOperation(EmailOperation operation, string address, Dictionary<string, bool> processed = null)
        {
            // create processed addresses dictionary if it does not exist
            if (processed == null)  processed = new Dictionary<string, bool>();

            // dont reprocess this address if we have already processed it
            if (processed.ContainsKey(address)) return null;

            // try to find a mailbox for the address on one of the tenants
            ExchangeService service = null;
            EmailAddress mailbox = null;
            foreach (WebCredentials cred in credentials)
            {
                service = new ExchangeService();
                service.Credentials = cred;
                service.Url = new Uri(ConfigurationManager.AppSettings["url"]);

                try
                {
                    NameResolutionCollection results = service.ResolveName("smtp:" + address);
                    if (results.Count > 0)
                    {
                        mailbox = results[0].Mailbox;
                        break;
                    }
                }
                catch (Exception e)
                {
                    return new ExchangeResult(address, StatusCode.Error, "Failed to resolve name: " + e.Message);
                }
            }

            // if we did not find a mailbox for the address on any of the tenants then report recipient not found
            if (mailbox == null) return new ExchangeResult(address, StatusCode.RecipientNotFound, "recipient not found");

            // add resolved address to processed list to prevent reprocessing
            processed.Add(mailbox.Address, true);

            // if this mailbox is a group/distribution list
            if (mailbox.MailboxType == MailboxType.PublicGroup)
            {
                // attempt to expand the group
                ExpandGroupResults group = null;
                try { group = service.ExpandGroup(mailbox.Address); }
                catch (Exception e)
                {
                    // report failure to expand group if an exception occurs during expansion
                    return new ExchangeResult(mailbox.Address, StatusCode.Error, "Failed to expand group: " + e.Message);
                }

                // for every member in the group
                ExchangeResult result = new ExchangeResult();
                foreach (EmailAddress member in group.Members)
                {
                    // recursively execute operation and log results
                    result.Log(ExecuteOperation(operation, member.Address, processed));
                }

                // return the results
                return result;
            }

            // if this is just a regular mailbox
            else if (mailbox.MailboxType == MailboxType.Mailbox)
            {
                // set impersonation to the mailbox address
                service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, mailbox.Address);

                // attempt to get some info to see if impersonation worked
                try { DateTime? dt = service.GetPasswordExpirationDate(mailbox.Address); }
                catch (Exception e)
                {
                    // if we were unable to impersonate the user then report error
                    return new ExchangeResult(mailbox.Address, StatusCode.Error, "impersonation failed: " + e.Message);
                }

                // delete email if operation is delete
                if (operation == EmailOperation.Delete)
                {
                    try
                    {
                        // find all instances of the email with message_id in the mailbox
                        FolderView folderView = new FolderView(int.MaxValue);
                        folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName);
                        folderView.Traversal = FolderTraversal.Shallow;
                        SearchFilter folderFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "AllItems");
                        FindFoldersResults folders = service.FindFolders(WellKnownFolderName.Root, folderFilter, folderView);
                        SearchFilter filter = new SearchFilter.IsEqualTo(EmailMessageSchema.InternetMessageId, message_id);
                        ItemView view = new ItemView(int.MaxValue);
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                        view.Traversal = ItemTraversal.Shallow;
                        List<ItemId> items = new List<ItemId>();
                        foreach (Item item in service.FindItems(folders.Folders[0].Id, filter, view))
                        {
                            items.Add(item.Id);
                        }

                        // if no instances of the email were found in the mailbox then report message_id not found
                        if (items.Count == 0)
                        {
                            return new ExchangeResult(mailbox.Address, StatusCode.MessageNotFound, "message_id not found");
                        }

                        // delete all found instances of the email with message_id from the mailbox
                        foreach (ServiceResponse response in service.DeleteItems(items, DeleteMode.SoftDelete, null, null))
                        {
                            // if we failed to delete an instance of the email then report an error
                            if (response.Result != ServiceResult.Success)
                            {
                                string message = "failed to delete email: " + response.ErrorCode + " " + response.ErrorMessage;
                                return new ExchangeResult(mailbox.Address, StatusCode.Error, message);
                            }
                        }
                    } catch (Exception e)
                    {
                        //report any errors we encounter
                        return new ExchangeResult(mailbox.Address, StatusCode.Error, "failed to delete email: " + e.Message);
                    }
                }

                // recover email if operation is recover
                else if (operation == EmailOperation.Restore)
                {
                    try
                    {
                        // find all instances of the email with message_id in the recoverable items folder
                        SearchFilter filter = new SearchFilter.IsEqualTo(EmailMessageSchema.InternetMessageId, message_id);
                        ItemView view = new ItemView(int.MaxValue);
                        view.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                        view.Traversal = ItemTraversal.Shallow;
                        List<ItemId> items = new List<ItemId>();
                        foreach (Item item in service.FindItems(WellKnownFolderName.RecoverableItemsDeletions, filter, view))
                        {
                            items.Add(item.Id);
                        }

                        // if no instances of the email with message_id were found in the recoverable items folder of the mailbox
                        if (items.Count == 0)
                        {
                            // report message_id not found
                            return new ExchangeResult(mailbox.Address, StatusCode.MessageNotFound, "message_id not found");
                        }

                        // move every instance of the email with message_id in the recoverable items folder to the inbox
                        foreach (ServiceResponse response in service.MoveItems(items, new FolderId(WellKnownFolderName.Inbox)))
                        {
                            // if we failed to move an instance of the email to the inbox then report an error
                            if (response.Result != ServiceResult.Success)
                            {
                                string message = "failed to recover email: " + response.ErrorCode + " " + response.ErrorMessage;
                                return new ExchangeResult(mailbox.Address, StatusCode.Error, message);
                            }
                        }
                    } catch (Exception e)
                    {
                        // report any errors we encounter
                        return new ExchangeResult(mailbox.Address, StatusCode.Error, "failed to recover email: " + e.Message);
                    }
                }

                // report successful operation
                return new ExchangeResult(mailbox.Address, StatusCode.Success, "success");
            }

            // report that the mailbox type is not one of the supported types
            return new ExchangeResult(mailbox.Address, StatusCode.Error, "Unsupported mailbox type: " + mailbox.MailboxType.ToString());
        }
    }
}
