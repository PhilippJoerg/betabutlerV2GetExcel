using System;
using System.Net.Http;
using System.Security.Cryptography;

namespace betabutlerV2GetExcel
{
    class WebUtil
    {
        public WebUtil() { }
        public string GetDocument(string container, string resourceName)
        {
            using (HttpClient httpClient = new HttpClient())
            {
                string resource = container + "/" + resourceName;
                string sasToken = GenerateStorageSasToken(resource, "thisIsTheStorageAccountURL", "thisIsTheStorageAccountKey");
                var response = httpClient.GetAsync(sasToken).Result;
                return (response.Content.ReadAsStringAsync().Result);
            }
        }

        public string GenerateStorageSasToken(string resourceName, string storageAccountUrl, string storageAccountKey)
        {
            var storageAccountName = storageAccountUrl.Remove(0, 8).Split('.')[0];
            var version = "2018-03-28";
            var startTime = DateTime.UtcNow;
            var startTimeIso = startTime.ToString("s") + "Z";
            var EndTimeIso = startTime.AddMinutes(10).ToString("s") + "Z";
            var hmacSha256 = new HMACSHA256 { Key = Convert.FromBase64String(storageAccountKey) };
            var payLoad = string.Format("{0}\n{1}\n{2}\n{3}\n{4}\n{5}\n{6}\n{7}\n\n\n\n\n", 
                "r",
                startTimeIso,
                EndTimeIso,
                "/blob/" + storageAccountName + "/" + resourceName,
                "",
                "",
                "https",
                "2018-03-28");
            var sasToken = storageAccountUrl + resourceName +
                    "?" +
                    "sp=r&st=" + startTimeIso + "&se=" + EndTimeIso + "&spr=https" +
                    "&sv=" + version +
                    "&sig=" + Uri.EscapeDataString(Convert.ToBase64String(hmacSha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(payLoad)))) +
                    "&sr=b";
            return sasToken;
        }
    }
}
