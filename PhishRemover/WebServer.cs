using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using CERTENROLLLib;
using System.Threading;

namespace PhishRemover
{
    public class PageResult
    {
        public int code = 200;
        public string message = "success";

        public PageResult(int code, string message)
        {
            this.code = code;
            this.message = message;
        }
    }

    public class WebServer
    {
        private HttpListener listener = new HttpListener();
        private volatile bool running = false;
        private int max_threads = 4;
        private string port = "3100";
        private volatile int num_threads = 0;
        private Dictionary<string, bool> client_ips = new Dictionary<string, bool>();
        public Dictionary<string, Func<string, PageResult>> pages = new Dictionary<string, Func<string, PageResult>>();

        public WebServer()
        {
            NameValueCollection config = ConfigurationManager.AppSettings;
            max_threads = Convert.ToInt32(config["max_threads"]);
            port = config["port"];
            string[] ips = config["client_ips"].Split(',');
            foreach (string ip in ips)
            {
                if (!client_ips.ContainsKey(ip))
                {
                    client_ips.Add(ip, true);
                }
            }
        }

        public bool Start()
        {
            try
            {
                running = true;
                listener = new HttpListener();
                InstallCertificate();
                listener.Prefixes.Add("https://*:" + port + "/");
                listener.Start();
                Thread dispatchThread = new Thread(DispatchRequests);
                dispatchThread.Priority = ThreadPriority.BelowNormal;
                dispatchThread.Start();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        public void Stop()
        {
            running = false;
            try { listener.Stop(); }
            catch { }
            listener.Close();
        }

        private void DispatchRequests()
        {
            while (running)
            {
                try
                {
                    HttpListenerContext context = listener.GetContext();
                    if (client_ips.ContainsKey(context.Request.RemoteEndPoint.ToString())) continue;
                    while (true)
                    {
                        if (num_threads < max_threads)
                        {
                            num_threads++;
                            Thread processThread = new Thread(() => ProcessRequests(context));
                            processThread.Priority = ThreadPriority.BelowNormal;
                            processThread.Start();
                            break;
                        }
                    }
                }
                catch (HttpListenerException) { break; }
                catch { }
                Thread.Sleep(1);
            }
        }

        private void ProcessRequests(HttpListenerContext context)
        {
            HttpListenerRequest request = context.Request;
            string json = "";
            using (StreamReader reader = new StreamReader(request.InputStream))
            {
                json = reader.ReadToEnd();
            }
            using (HttpListenerResponse response = context.Response)
            {
                PageResult result;
                if (pages.ContainsKey(request.Url.AbsolutePath))
                {
                    try
                    {
                        result = pages[request.Url.AbsolutePath](json);
                    }
                    catch (Exception e)
                    {
                        result = new PageResult(500, e.Message);
                    }
                }
                else
                {
                    result = new PageResult(404, "page not found");
                }

                response.StatusCode = result.code;
                using (StreamWriter writer = new StreamWriter(response.OutputStream))
                {
                    writer.Write(result.message);
                }
                num_threads--;
            }
        }

        private void InstallCertificate()
        {
            //create cert if it does not exist
            X509Store store = new X509Store(StoreName.TrustedPeople, StoreLocation.LocalMachine);
            store.Open(OpenFlags.ReadWrite);
            X509Certificate2Collection certs = store.Certificates.Find(X509FindType.FindBySubjectName, "PhishRemover", false);
            if (certs.Count == 0) store.Add(CreateSelfSignedCertificate("PhishRemover"));
            string cert_hash = certs[0].GetCertHashString();
            store.Close();

            //bind certificate to ssl port
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C netsh http add sslcert ipport=0.0.0.0:" + port + " certhash=" + cert_hash + " appid={6232e5bc-ebd6-4dbb-98b5-7e65ed513236}";
            process.StartInfo = startInfo;
            process.Start();
        }

        private X509Certificate2 CreateSelfSignedCertificate(string subjectName)
        {
            // create DN for subject and issuer
            var dn = new CX500DistinguishedName();
            dn.Encode("CN=" + subjectName, X500NameFlags.XCN_CERT_NAME_STR_NONE);

            // create a new private key for the certificate
            CX509PrivateKey privateKey = new CX509PrivateKey();
            privateKey.ProviderName = "Microsoft Base Cryptographic Provider v1.0";
            privateKey.MachineContext = true;
            privateKey.Length = 2048;
            privateKey.KeySpec = X509KeySpec.XCN_AT_SIGNATURE; // use is not limited
            privateKey.ExportPolicy = X509PrivateKeyExportFlags.XCN_NCRYPT_ALLOW_PLAINTEXT_EXPORT_FLAG;
            privateKey.Create();

            // Use the stronger SHA512 hashing algorithm
            var hashobj = new CObjectId();
            hashobj.InitializeFromAlgorithmName(ObjectIdGroupId.XCN_CRYPT_HASH_ALG_OID_GROUP_ID,
                ObjectIdPublicKeyFlags.XCN_CRYPT_OID_INFO_PUBKEY_ANY,
                AlgorithmFlags.AlgorithmFlagsNone, "SHA512");

            // add extended key usage if you want - look at MSDN for a list of possible OIDs
            var oid = new CObjectId();
            oid.InitializeFromValue("1.3.6.1.5.5.7.3.1"); // SSL server
            var oidlist = new CObjectIds();
            oidlist.Add(oid);
            var eku = new CX509ExtensionEnhancedKeyUsage();
            eku.InitializeEncode(oidlist);

            // Create the self signing request
            var cert = new CX509CertificateRequestCertificate();
            cert.InitializeFromPrivateKey(X509CertificateEnrollmentContext.ContextMachine, privateKey, "");
            cert.Subject = dn;
            cert.Issuer = dn; // the issuer and the subject are the same
            cert.NotBefore = DateTime.Now;
            // this cert expires immediately. Change to whatever makes sense for you
            cert.NotAfter = DateTime.Now;
            cert.X509Extensions.Add((CX509Extension)eku); // add the EKU
            cert.HashAlgorithm = hashobj; // Specify the hashing algorithm
            cert.Encode(); // encode the certificate

            // Do the final enrollment process
            var enroll = new CX509Enrollment();
            enroll.InitializeFromRequest(cert); // load the certificate
            enroll.CertificateFriendlyName = subjectName; // Optional: add a friendly name
            string csr = enroll.CreateRequest(); // Output the request in base64
            // and install it back as the response
            enroll.InstallResponse(InstallResponseRestrictionFlags.AllowUntrustedCertificate,
                csr, EncodingType.XCN_CRYPT_STRING_BASE64, ""); // no password
            // output a base64 encoded PKCS#12 so we can import it back to the .Net security classes
            var base64encoded = enroll.CreatePFX("", // no password, this is for internal consumption
                PFXExportOptions.PFXExportChainWithRoot);

            // instantiate the target class with the PKCS#12 data (and the empty password)
            return new X509Certificate2(
                Convert.FromBase64String(base64encoded), "",
                // mark the private key as exportable (this is usually what you want to do)
                X509KeyStorageFlags.Exportable
            );
        }
    }
}
