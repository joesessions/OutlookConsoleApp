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
        private const string ClientId = "93e33143-ebea-4267-b9f1-5fec63ab5bd9";
        private const string AadInstance = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
        private const string Tenant = "JoeAsDeveloperhotmail.onmicrosoft.com";
        private const string Resource = "https://microsoft.graph.net";
        private const string ClientSecret = "Go.74gmOpfhiQk8-agE-R0ZDFwf9O.KOE_";
        private const string Redirect_uri = "https%3A%2F%2FJoeAsDeveloper.onmicrosoft.com%2FAppForAzureAccess";
        private const string Redirect_uri_Plain = "https://JoeAsDeveloper.onmicrosoft.com/AppForAzureAccess";
        private const string Scope = "https%3A%2F%2Fgraph.microsoft.com%2F.default";
        private static HttpClient _HttpClient = new HttpClient();

        private string AuthToken{ get; set; }

        private const string GetThisUserToken = "EwBgA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAVTjC/Fx2kluU3lShq30pvDQdB6a9Ef1DUBVrFWiRB8HtKtF15A4bWt3/iKb7PjH4/4AEHsHeBu62Ue6051/R1ZzqayTiA3dGddjRWT7XJtaEUItaFd6MnEV9eGQ0vtWHprutL19DrYihvbzE6jIPBjWQlxIhXiHuCEPZbLSMjEZmt0f0srec4s8402+rcD0+GYfQSPIPqX/OZx+YdPfO5/YLRO3MWB9+S34XPGBlbtgIGCz7MNZAE37e/4ah6S6r8soKH9zn1k321HQ4/1Kn9z5Ak1FWNspo+cQhhJVct4kORzQXhSW95VxIR+dp4xBjQuGD/mexKtvAbmWEMRhEjUDZgAACIUvq/HPVw/iMAJTBqKpCY8FdHT+L+/73nN2Uph+bPulerh1gy8BxNXkkvvlWpc9JU1fGfTLZhVmJG7NlsMbsRlkotVjkmOWcLRqQNf7rnQv+1NI3889M25Y8FlKSdER9O2JDB8hTsQodd19yRwtiaULzW7KPA8DXdESQ/tCMEWBXSesol3q2esQtzURCvASuEHklraobYl3uFXvAXFawOxV9Xbzlzil9RUH/8481lSz/sa9u21c9cwwzvOGYR/XxCbyFxrXkES5EJRR4ihWqtQtJ629W2M1lkLCvB7Fqk7e8rbcbnLxWsTxMLDZHhYzGMQEhwdO48Exjcm+k2Vh9jS/EZ3UxkiibMddPx8P/ks/DwFgux1Ae7X/8mP3Uuzts4EJjOLPwMMbZ7b7IlpNIYd+BRQjhxYbFnOh7tt9s4qj65qlqK4TvbHPVRLMzbs0vxofrjiqgOjo3AE/5U5gyr5vYuH4M6zGSDb4Q+PD7zugqq6CJ9YhYpbr89uNzfB1SPxF4aDgK03PwjcDIjSanEXaL6uesi8reedZFb6sZlDsjj4cH9YXv8SaOcNnuYknUEgCM76C0TH8L1Bra/48KPy3JbSE5oxKC8c0w9J36zjwayWkM9mZAzJvf26rrLW1DH/DZ+blybV+EtYAZ+tOunDmBpild+6mxtWtyowzCffTC6eEaIyOoMVFO4/oOiCrb6D3c4ZPqi6OrcD5E+J0sCQmojUDtlaAqqibB3YvGWtd/ulXmPmahPa3eXQC";
        private const string ReadWriteEmailToken = "EwBoA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAASXFP8pCGIiff4W1sSl6XWDBBmWRNxCJcpVEqp2aS02L8B0WChztZekP0KedN6T7Vktr9y21Y7sHE/wWxV5InYFWpXGn9YOofCPs/9Vv8dMUYvKx5s+KMsNzI2vtOOCYKFVUh8aScWUQAv1c7+dJ1NQcp4HARO5oHs/s29ji/oxciVdEXv37lK7PqIxSLFcXzzsB80y/rcmSvoKWIndHDO/jZ0pCOaYxzzCDs77qnyW+lJQ9jbBs+o8ok4FumTFgrt1BEzvf+tu0C8IllmCBI+t6q2XdLKPo67REMC5MEMdOHhY/WRugXyRmCMbAOuRfcbyJuB2UXeRNO8xrC9ZPAk4DZgAACAKNtK+WbKCnOAJiZyMWVtKJY1zLYr4GosS1uC8tHcrL2SM+T+0clPM0BitQNE2Mr/Prnjt2xE8fqLOnhLNlY5ubEH5PNRV5iR7V+jnl2foWZw+0p7LGVvxivD/ARQrAa9hYgkAwnqRo5nZgy3glLExj2ygC99YHkxG4jtMUi7awqg4UlGLauMjTm5rbkKbqwWljs9DI9CErM8/3rNfQnL/p0/zSkSPWO7+ruYMFLiA2mN2V8493glnBbQMS/aJjM7EdXL76m2ful6z1AC8yiE4LgAXvWHQHkCl+I+MS5Wx8d9rCP8XAIE3bm21mp/xU1nNBHP5voQKY/w2amy7ZAgJKLiY8ydBzyFWX5ft48ctwORTh4Ur3uasrPRZx/8rGtlVH22nyOJLXfphcAg2mNkvHVw2ZTwuctLctRKTecUqu5BE7s96/5xS5I88uTcTDt8sEMglUKtTwgA8oFsJ2HhSnvsjbN+P4vNY4qBSwbrp23pXvw/HrbKqKAeqOVakjX0F8P9jduuKXh3Htkx2b0jMyEt6WXvwu4dFidlnjBY/K6mAqrg6/sKnPQVIAqSNN/8HIz9w2Sde9sWjVz97aQ80e8eahVGHg2N/9vBeANCDrzMHzyzKH5l2HSPSvahU5YvJdFoFy7FjUy+a3/2/ScaDv7lAAxjobHGi9yW3GhvNlWB6VDvYglv6QnXKrLOHMPwlgDfo0BH2pi0pVXvymyE4ElfA6IeUPbHgbW9NilJV010G5MpqkFGfBVxJ3O+s4NMWoeAI=";
        private const string ClientCredentialsToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlBlMUYwOGNWVmQ3ckdkZ2c2bUhmam82NFFMcTh1bU9BWEdyeDlrblJOUEUiLCJhbGciOiJSUzI1NiIsIng1dCI6IlNzWnNCTmhaY0YzUTlTNHRycFFCVEJ5TlJSSSIsImtpZCI6IlNzWnNCTmhaY0YzUTlTNHRycFFCVEJ5TlJSSSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDcvIiwiaWF0IjoxNTkxNTQ2Nzg5LCJuYmYiOjE1OTE1NDY3ODksImV4cCI6MTU5MTYzMzQ4OSwiYWlvIjoiNDJkZ1lEaGdlZlRONGF3bXZrdDl6ei9acWppZUFBQT0iLCJhcHBfZGlzcGxheW5hbWUiOiJBcHBGb3JBenVyZUFjY2VzcyIsImFwcGlkIjoiOTNlMzMxNDMtZWJlYS00MjY3LWI5ZjEtNWZlYzYzYWI1YmQ5IiwiYXBwaWRhY3IiOiIxIiwiaWRwIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3LyIsInJoIjoiMC5BUm9BdjRqNWN2R0dyMEdScXkxODBCSGJSME14NDVQcTYyZEN1ZkZmN0dPclc5a2FBQUEuIiwidGVuYW50X3JlZ2lvbl9zY29wZSI6IldXIiwidGlkIjoiNzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3IiwidXRpIjoiaWNwNGViNms5MEtwcERGcFM5NHdBQSIsInZlciI6IjEuMCIsInhtc190Y2R0IjoxMjg5MjQxNTQ3fQ.MoOFkPRHYlzTPII38k8amueWcKshzlE_CVx6QwWRFd3J1fsaRehf14fTdhf1xJkF-O_REni8LN_Bp83AZ0u3FJMdYsdjLkpFYlo3KbLyVwJH9j55YxLPhY4ksjvgTt72cb71AQu_Kjcjn3qg968wr4UPIqA1mtrDZfChf9W9i1Hlb1yI_x5Q4nn40uZ3neFDgWkN8skWxo3dY8OIdz0n_GEu6Ip5CNm6Wr1HpT4svWUA_QM2vxb_T8HmaDYmCtkLec9TXSD8O4ejnPVWfqV_jfrWe_DgICSSSC6PPyvjP5BIXcaPuKjaHw7kd5hXlDeFeEzLFSsy-AbD0kDhi40OXA";


        public async Task<string> GetAuthToken2()
        {
            string uri = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&state="
                        + "&client_id=" + ClientId + "&scope=" + Scope
                        + "&redirect_uri=" + Redirect_uri;
            var getResult = await _HttpClient.GetAsync(uri);
            string firstResponse = "";
            if (getResult.Content != null)
            {
                firstResponse = await getResult.Content.ReadAsStringAsync();
            }

            string uaid = FindValueInFirstResponse(firstResponse, "uaid");
            string msproxy = FindValueInFirstResponse(firstResponse, "msproxy");
            string issuer = FindValueInFirstResponse(firstResponse, "issuer");
            string tenant = FindValueInFirstResponse(firstResponse, "tenant");
            string ui_locales = FindValueInFirstResponse(firstResponse, "ui_locales");

            string uri2 = $"https://login.live.com/oauth20_authorize.srf?client_id={ClientId}"
                        + $"&scope={Scope}&redirect_uri={Redirect_uri}"
                        + $"&response_type=code&uaid={uaid}&msproxy={msproxy}"
                        + $"$issuer={issuer}$tenant={tenant}$ui_locales={ui_locales}";
            var getResult2 = await _HttpClient.GetAsync(uri);

            HttpResponseHeaders secondResult = null;
            if (getResult2.Content != null)
            {
                secondResult = getResult2.Headers;
            }

            //find indexof 
            //make substring of that
            //find first next index of \
            //make substring of that and assign it to uaid

            string key2 = "\"sessionId\":\"";
            string key = "\"correlationId\":\"";
            int startIndex = firstResponse.IndexOf(key);
            string frontChoppedOff = firstResponse.Substring(startIndex + key.Length);
            string authorization_code = frontChoppedOff.Substring(0, 36);

            var postData = new List<KeyValuePair<string, string>>();
            postData.Add(new KeyValuePair<string, string>("grant_type", "authorization_code"));
            postData.Add(new KeyValuePair<string, string>("code", authorization_code));
            postData.Add(new KeyValuePair<string, string>("redirect_uri", Redirect_uri_Plain));
            postData.Add(new KeyValuePair<string, string>("client_id", ClientId));
            postData.Add(new KeyValuePair<string, string>("client_secret", ClientSecret));
            HttpContent content = new FormUrlEncodedContent(postData);
            content.Headers.Remove("Content-Type");
            content.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            var postResult = _HttpClient.PostAsync(AadInstance, content);
            string token = "";
            if (postResult.Result.Content != null)
            {
                token = await postResult.Result.Content.ReadAsStringAsync();
            }
            
            var goldenTicket =  token;

            return "yomama";






        }
        
        private string FindValueInFirstResponse(string firstResponse, string key)
        {
            int startIndex = firstResponse.IndexOf(key);
            string frontChoppedOff = firstResponse.Substring(startIndex + key.Length + 1);
            int strLength = frontChoppedOff.IndexOf(@"\");
            string ret = frontChoppedOff.Substring(0, strLength);
            return ret;
        }

        public async Task<string> GetThisUser()
        {
            string users = null;
            var uri = "https://graph.microsoft.com/v1.0/me/";

            _HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", GetThisUserToken);
            var getResult = await _HttpClient.GetAsync(uri);

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

            _HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ReadWriteEmailToken);
            var getResult = await _HttpClient.GetAsync(uri);

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

            _HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ReadWriteEmailToken);

            var getResult = await _HttpClient.PostAsync(uri, data);

            return true;
        }

    }

    public class Sendable
    {
        public Message Message { get; set; }
        public bool saveToSentItems { get; set; }
    }


}
