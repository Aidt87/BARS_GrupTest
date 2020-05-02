using System.Collections.Generic;

namespace BARS_GrupTest
{
    internal class Installed
    {
        public string client_id { get; set; }
        public string project_id { get; set; }
        public string auth_uri { get; set; } = "https://accounts.google.com/o/oauth2/auth";
        public string token_uri { get; set; } = "https://oauth2.googleapis.com/token";
        public string auth_provider_x509_cert_url { get; set; } = "https://www.googleapis.com/oauth2/v1/certs";
        public string client_secret { get; set; }
        public List<string> redirect_uris { get; set; } = new List<string>(new string[] { "urn:ietf:wg:oauth:2.0:oob", "http://localhost" });
    }

    internal class GoogleAuthСredentials
    {
        public Installed installed { get; set; }
    }
}
