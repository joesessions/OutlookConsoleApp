using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace OutlookConsole
{
    class GraphClient
    {
        private const string ClientId = "api://93e33143-ebea-4267-b9f1-5fec63ab5bd9";
        private const string AadInstance = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        private const string Tenant = "JoeAsDeveloperhotmail.onmicrosoft.com";
        private const string Resource = "https://microsoft.graph.net";
        private const string ClientSecret = "Go.74gmOpfhiQk8-agE-R0ZDFwf9O.KOE_";

        private static HttpClient HttpClient = new HttpClient();
      
        private const string GetThisUserToken = "EwBgA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAVTjC/Fx2kluU3lShq30pvDQdB6a9Ef1DUBVrFWiRB8HtKtF15A4bWt3/iKb7PjH4/4AEHsHeBu62Ue6051/R1ZzqayTiA3dGddjRWT7XJtaEUItaFd6MnEV9eGQ0vtWHprutL19DrYihvbzE6jIPBjWQlxIhXiHuCEPZbLSMjEZmt0f0srec4s8402+rcD0+GYfQSPIPqX/OZx+YdPfO5/YLRO3MWB9+S34XPGBlbtgIGCz7MNZAE37e/4ah6S6r8soKH9zn1k321HQ4/1Kn9z5Ak1FWNspo+cQhhJVct4kORzQXhSW95VxIR+dp4xBjQuGD/mexKtvAbmWEMRhEjUDZgAACIUvq/HPVw/iMAJTBqKpCY8FdHT+L+/73nN2Uph+bPulerh1gy8BxNXkkvvlWpc9JU1fGfTLZhVmJG7NlsMbsRlkotVjkmOWcLRqQNf7rnQv+1NI3889M25Y8FlKSdER9O2JDB8hTsQodd19yRwtiaULzW7KPA8DXdESQ/tCMEWBXSesol3q2esQtzURCvASuEHklraobYl3uFXvAXFawOxV9Xbzlzil9RUH/8481lSz/sa9u21c9cwwzvOGYR/XxCbyFxrXkES5EJRR4ihWqtQtJ629W2M1lkLCvB7Fqk7e8rbcbnLxWsTxMLDZHhYzGMQEhwdO48Exjcm+k2Vh9jS/EZ3UxkiibMddPx8P/ks/DwFgux1Ae7X/8mP3Uuzts4EJjOLPwMMbZ7b7IlpNIYd+BRQjhxYbFnOh7tt9s4qj65qlqK4TvbHPVRLMzbs0vxofrjiqgOjo3AE/5U5gyr5vYuH4M6zGSDb4Q+PD7zugqq6CJ9YhYpbr89uNzfB1SPxF4aDgK03PwjcDIjSanEXaL6uesi8reedZFb6sZlDsjj4cH9YXv8SaOcNnuYknUEgCM76C0TH8L1Bra/48KPy3JbSE5oxKC8c0w9J36zjwayWkM9mZAzJvf26rrLW1DH/DZ+blybV+EtYAZ+tOunDmBpild+6mxtWtyowzCffTC6eEaIyOoMVFO4/oOiCrb6D3c4ZPqi6OrcD5E+J0sCQmojUDtlaAqqibB3YvGWtd/ulXmPmahPa3eXQC";
        private const string ReadWriteEmailToken = "EwBoA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAATg5ackd5FfieqBsm2vuEsNgl0LRUO9IGsx/nSKx2YCkVveanN4QZwmB2FhwEXhi0wvwfshhMpstonxYpffWolLIV36D9a8DxIETLcMWPcZB/coKSr7ULCB1jHilJBu1naT4Lf8dt32k+3/8jiAu3rUGYdv+WhzBJKKb1Ia0u+nt7r5Dk7VDpCUCOBfR7brxjvetw4zCJcyvTwvCajiN7PwBxeWg0GFw0z8SBGb/7CEVgN8z3qWHAE47+g6nsrmAaxz6BMwVTy/e/gh0VD7fvXGg6Nli/IWG96hvVe61uRRVdz5S4w/YUGCEGi9KwArQTPb6uTQDPgO5jdfS9FyTxAADZgAACAK3ZyLTTAtuOAIMoqtZXmWt2K+AtOFYsdXS19Nlo0bB0tYGkMxBtplmPwyzRoBTU7DX7lKeG2FpqcVTLFLr1+DLfywRsnvxrLNFX0KVVeUvVpZ2/bLlpVD7I/0H3eVSUA6ErdV61WBwfJV4mUuInzqVzE2wHKcTUSOlkOEw/jXxHFWyf/mEjem4iQxCAhYdLdf2WdpUeT9ye6y5xTC/mpUXmgQonmEGGU8vJb0rWf1uLtEmBBVNNFx4phWyBELLyyCW7oizw1/MZ0GLbAJ3KhwUHuI3LwFYBcb1ZFBS/hJibLduU+fFXUT6prv2pdcQpsfuJ5Hv+8lcRCziyIQ9+KO9ahXKiM79keDHcenmMlkZOiuP/PnPywN1Vq8BTl1Yz2iRYKqfsqFzZtl/wXD+uH1xo0x/539Foq/aVcI3PQQnoPKPmybHPZosUYfo66jJxggouXatqSx0ewTvbyTlBU4I5bYg7x7mKrf30n89ZqVyAz4yWkYeDC7KvBA0ydMPkCUSpv27V/i686BUvV+roy91aFv3NIGWKCc1FxjvJ0aXaikYPtzT2HKwO3sSEToeknv7F4GBrG6J0NNpA3AS+x9A2b75UrTIVgDtOvHekW8fxEflIjY7o/vbxPcjXm5KhUtfw0U+dss/gLx9rhqfWZHqKWOL+xWbmkgayeTj5eQdQkJjPZCfKCq8AwjeSHl3fQLRdxaLUFbWnl9GUABpVTXRfYbYAtFoGZT6jlud20yc4xUzH1UOegVA609Q6oC0Mqr6eAI=";

        public async Task<string> GetThisUser()
        {
            string users = null;
            var uri = "https://graph.microsoft.com/v1.0/me/";

            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", GetThisUserToken);
            var getResult = await HttpClient.GetAsync(uri);

            if (getResult.Content != null)
            {
                users = await getResult.Content.ReadAsStringAsync();
            }

            return users;
        }

        public async Task<List<SimpleEmail>> GetTenEmails()
        {
            string emailsJsonString = null;
            var uri = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages";

            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ReadWriteEmailToken);
            var getResult = await HttpClient.GetAsync(uri);

            if (getResult.Content != null)
            {
                emailsJsonString = await getResult.Content.ReadAsStringAsync();
            }

            var value = JsonConvert.DeserializeObject<RootObject>(emailsJsonString);

            var retEmails = new List<SimpleEmail>();

            foreach (Value val in value.value)
            {
                retEmails.Add(
                    new SimpleEmail() {
                        Subject = val.subject,
                        FromAddress = val.from.emailAddress.Address,
                        Date = val.sentDateTime,
                        Id = val.id,
                        Body = HtmlToPlainText(val.body.content) // Remove the method to get more HTML artifacts.
                    }
                );
            }
                       
            return retEmails;
        }

        public class Value
        {
            [JsonProperty("@odata.etag")]
            public string etag { get; set; }
            public string accountnumber { get; set; }
            public string accountid { get; set; }
            public string subject { get; set; }
            public From from { get; set; }
            public DateTime sentDateTime { get; set; }
            public string id { get; set; }
            public Body body { get; set; }

        }

        public class Body
        {
            public string contentType { get; set; }
            public string content { get; set; }
        }

        public class From
        {
            [JsonProperty("emailAddress")]
            public EmailAddress emailAddress { get; set; }
        }

        public class RootObject
        {
            [JsonProperty("@odata.context")]
            public string context { get; set; }
            public List<Value> value { get; set; }
        }

        public class SimpleEmail
        {
            public string Subject { get; set; }
            public string FromAddress { get; set; }
            public DateTime Date { get; set; }
            public string Id { get; set; }
            public string Body { get; set; }
        }

        private static string HtmlToPlainText(string html)
        {
            const string tagWhiteSpace = @"(>|$)(\W|\n|\r)+<";//matches one or more (white space or line breaks) between '>' and '<'
            const string stripFormatting = @"<[^>]*(>|$)";//match any character between '<' and '>', even when end tag is missing
            const string lineBreak = @"<(br|BR)\s{0,1}\/{0,1}>";//matches: <br>,<br/>,<br />,<BR>,<BR/>,<BR />
            var lineBreakRegex = new Regex(lineBreak, RegexOptions.Multiline);
            var stripFormattingRegex = new Regex(stripFormatting, RegexOptions.Multiline);
            var tagWhiteSpaceRegex = new Regex(tagWhiteSpace, RegexOptions.Multiline);

            var text = html;
            //Decode html specific characters
            text = System.Net.WebUtility.HtmlDecode(text);
            //Remove tag whitespace/line breaks
            text = tagWhiteSpaceRegex.Replace(text, "><");
            //Replace <br /> with line breaks
            text = lineBreakRegex.Replace(text, Environment.NewLine);
            //Strip formatting
            text = stripFormattingRegex.Replace(text, string.Empty);

            return text;
        }

        public async Task<bool> SendEmail()  // string toEmail, string subject, string body)
        {
            var graphClient = new GraphClient();

            var message = new Message
            {
                Subject = "Hello Test Email",// subject,
                Importance = Importance.Low,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Hello World" //body
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "joesessions@gmail.com"// toEmail
                        }
                    }
                }
            };
            
            var uri = "https://graph.microsoft.com/v1.0/me/sendMail";
            var sendable = new Sendable()
            {
                Message = message,
                //saveToSentItems = false
            };
            var json = JsonConvert.SerializeObject(sendable);
            var data = new StringContent(json, Encoding.UTF8, "application/json");

            HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ReadWriteEmailToken);

            var getResult = await HttpClient.PostAsync(uri, data);

            return true;
        }

    }

    public class Sendable
    {
        public Message Message { get; set; }
        public bool saveToSentItems { get; set; }
    }


}
