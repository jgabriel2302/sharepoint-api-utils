
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

public class SPRestApi
{
    private readonly string siteUrl;
    private readonly string accessToken;
    private readonly HttpClient httpClient;

    public SPRestApi(string siteUrl, string accessToken)
    {
        this.siteUrl = siteUrl;
        this.accessToken = accessToken;
        this.httpClient = new HttpClient();
        this.httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        this.httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    private string BuildListUrl(string listName, string endpoint = "")
    {
        return $"{siteUrl}/_api/web/lists/getbytitle('{listName}'){endpoint}";
    }

    public async Task<string> GetItemsAsync(string listName)
    {
        var url = BuildListUrl(listName, "/items");
        var response = await httpClient.GetAsync(url);
        return await response.Content.ReadAsStringAsync();
    }

    public async Task<string> GetUserInfoAsync()
    {
        var url = $"{siteUrl}/_api/web/currentuser";
        var response = await httpClient.GetAsync(url);
        return await response.Content.ReadAsStringAsync();
    }

    public async Task<string> AddItemAsync(string listName, object itemData)
    {
        var url = BuildListUrl(listName, "/items");
        var payload = JsonConvert.SerializeObject(itemData);
        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        var response = await httpClient.PostAsync(url, content);
        return await response.Content.ReadAsStringAsync();
    }

    public async Task<string> UpdateItemAsync(string listName, int itemId, object itemData)
    {
        var url = BuildListUrl(listName, $"/items({itemId})");
        var payload = JsonConvert.SerializeObject(itemData);
        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        var request = new HttpRequestMessage(new HttpMethod("MERGE"), url)
        {
            Content = content
        };
        request.Headers.Add("IF-MATCH", "*");
        var response = await httpClient.SendAsync(request);
        return await response.Content.ReadAsStringAsync();
    }

    public async Task<bool> DeleteItemAsync(string listName, int itemId)
    {
        var url = BuildListUrl(listName, $"/items({itemId})");
        var request = new HttpRequestMessage(HttpMethod.Delete, url);
        request.Headers.Add("IF-MATCH", "*");
        var response = await httpClient.SendAsync(request);
        return response.IsSuccessStatusCode;
    }

    public async Task<string> GetAccessTokenAsync()
    {
        using (var client = new HttpClient())
        {
            var tokenEndpoint = $"https://accounts.accesscontrol.windows.net/{tenantId}/tokens/OAuth/2";
            var resource = $"{siteUrl}@{tenantId}";
            var postData = new Dictionary<string, string>
            {
                { "grant_type", "client_credentials" },
                { "client_id", $"{clientId}@{tenantId}" },
                { "client_secret", clientSecret },
                { "resource", resource }
            };
    
            var content = new FormUrlEncodedContent(postData);
            var response = await client.PostAsync(tokenEndpoint, content);
            var responseContent = await response.Content.ReadAsStringAsync();
    
            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Erro ao obter token: {response.StatusCode} - {responseContent}");
            }
    
            var jsonDoc = JsonDocument.Parse(responseContent);
            var accessToken = jsonDoc.RootElement.GetProperty("access_token").GetString();
            return accessToken;
        }
    }

}
