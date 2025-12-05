using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Net.Http.Headers;

public class SPRestApi
{
    // =========================
    // Utils
    // =========================
    public static class Utils
    {
        public static byte[] Base64ParaBytes(string base64, string tipo = "image/jpeg")
        {
            // Suporta "data:image/jpeg;base64,AAAA" ou só "AAAA"
            var partes = base64.Split(',');
            var dados = partes.Length > 1 ? partes[1] : partes[0];
            return Convert.FromBase64String(dados);
        }

        public static byte[] Base64ParaArrayBuffer(string base64)
        {
            var partes = base64.Split(',');
            var dados = partes.Length > 1 ? partes[1] : partes[0];
            return Convert.FromBase64String(dados);
        }
    }

    private readonly HttpClient _httpClient;

    public string Site { get; private set; }
    public string ListaAtual { get; private set; }
    public string Type { get; private set; }

    /// <summary>
    /// Construtor principal.
    /// </summary>
    /// <param name="site">URL do site SharePoint</param>
    /// <param name="lista">Nome da lista padrão (opcional)</param>
    /// <param name="bearerToken">Token de autenticação (Authorization: Bearer ...)</param>
    public SPRestApi(
        string site = "https://<seu contoso>.sharepoint.com/sites/<seu site>",
        string lista = null,
        string bearerToken = null)
    {
        Site = site.TrimEnd('/');
        ListaAtual = lista;
        Type = EncodeEntityType(ListaAtual);

        _httpClient = new HttpClient();

        // Se você usar AAD: passe o token aqui.
        if (!string.IsNullOrEmpty(bearerToken))
        {
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", bearerToken);
        }
    }

    public void SetLista(string listaName)
    {
        ListaAtual = listaName;
    }

    public SPRestApi SetListType(string listType)
    {
        Type = listType;
        return this;
    }

    public SPRestApi GetLista(string listaName)
    {
        return new SPRestApi(Site, listaName);
    }

    public string EncodeEntityType(string lista, bool useDefaultType = false)
    {
        if (useDefaultType && !string.IsNullOrEmpty(Type))
            return Type;

        var nome = (lista ?? string.Empty)
            .Replace(" ", "_x0020_")
            .Replace("_", "_x005f_");

        return $"SP.Data.{nome}ListItem";
    }

    public string BuildListUrl(string lista, string endpoint = "")
    {
        if (string.IsNullOrEmpty(lista))
            throw new InvalidOperationException("Lista não definida.");

        return $"{Site}/_api/web/lists/getbytitle('{lista}'){endpoint}";
    }

    public string BuildSharePointUrl(string endpoint = "")
    {
        return $"{Site}/_api/web/{endpoint}";
    }

    // =========================
    // Método genérico de request
    // =========================
    private async Task<JsonDocument> RequestAsync(
        string url,
        HttpMethod method,
        IDictionary<string, string> headers = null,
        HttpContent body = null)
    {
        var request = new HttpRequestMessage(method, url);

        if (headers != null)
        {
            foreach (var kvp in headers)
            {
                request.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value);
            }
        }

        if (body != null)
        {
            request.Content = body;
        }

        var response = await _httpClient.SendAsync(request);
        var str = await response.Content.ReadAsStringAsync();

        response.EnsureSuccessStatusCode();

        return JsonDocument.Parse(str);
    }

    // =========================
    // CRUD de itens
    // =========================

    public async Task<JsonDocument> AddItemAsync(
        Dictionary<string, object> data,
        string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista, "/items");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"},
            {"Content-Type", "application/json;odata=verbose"}
            // Não usamos X-RequestDigest aqui (autenticação moderna)
        };

        var payloadDict = new Dictionary<string, object>(data ?? new Dictionary<string, object>())
        {
            ["__metadata"] = new Dictionary<string, object>
            {
                ["type"] = EncodeEntityType(lista, lista == ListaAtual)
            }
        };

        // Se quiser forçar Title sempre presente:
        if (!payloadDict.ContainsKey("Title"))
            payloadDict["Title"] = "";

        var json = JsonSerializer.Serialize(payloadDict);
        var content = new StringContent(json, Encoding.UTF8, "application/json");

        return await RequestAsync(url, HttpMethod.Post, headers, content);
    }

    public Task<JsonDocument> InsertItemAsync(
        Dictionary<string, object> data,
        string lista = null)
    {
        return AddItemAsync(data, lista);
    }

    public async Task<JsonDocument> UpdateItemAsync(
        int id,
        Dictionary<string, object> data,
        string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista, $"/items({id})");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"},
            {"Content-Type", "application/json;odata=verbose"},
            {"IF-MATCH", "*"},
            {"X-HTTP-Method", "MERGE"}
            // Sem X-RequestDigest
        };

        var payloadDict = new Dictionary<string, object>(data ?? new Dictionary<string, object>())
        {
            ["__metadata"] = new Dictionary<string, object>
            {
                ["type"] = EncodeEntityType(lista, lista == ListaAtual)
            }
        };

        var json = JsonSerializer.Serialize(payloadDict);
        var content = new StringContent(json, Encoding.UTF8, "application/json");

        // MERGE em SharePoint normalmente é POST + X-HTTP-Method
        return await RequestAsync(url, HttpMethod.Post, headers, content);
    }

    public async Task<bool> DeleteItemAsync(int id, string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista, $"/items({id})");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"},
            {"IF-MATCH", "*"},
            {"X-HTTP-Method", "DELETE"}
        };

        var doc = await RequestAsync(url, HttpMethod.Post, headers);
        // Se chegou aqui, EnsureSuccessStatusCode passou
        return true;
    }

    public async Task<JsonDocument> GetItemsAsync(
        Dictionary<string, string> @params = null,
        string lista = null)
    {
        lista ??= ListaAtual;

        var baseUrl = BuildListUrl(lista, "/items");
        var urlBuilder = new StringBuilder(baseUrl);

        if (@params != null && @params.Count > 0)
        {
            urlBuilder.Append("?");
            bool first = true;
            foreach (var kvp in @params)
            {
                if (!first) urlBuilder.Append("&");
                first = false;
                urlBuilder.Append($"${kvp.Key}={Uri.EscapeDataString(kvp.Value)}");
            }
        }

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };

        return await RequestAsync(urlBuilder.ToString(), HttpMethod.Get, headers);
    }

    public async Task<JsonDocument> GetItemByIdAsync(
        int id,
        string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista, $"/items({id})");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };

        return await RequestAsync(url, HttpMethod.Get, headers);
    }

    // =========================
    // Metadados, user e site
    // =========================

    public async Task<JsonDocument> GetListMetadataAsync(string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista, "/fields?$select=Id,EntityPropertyName,Choices,Title,TypeAsString");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };

        return await RequestAsync(url, HttpMethod.Get, headers);
    }

    public async Task<JsonDocument> GetFieldMetadataByNameAsync(
        string fieldName,
        string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista,
            $"/fields/getbytitle('{fieldName}')?$select=Id,EntityPropertyName,Choices,Title,TypeAsString");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };

        return await RequestAsync(url, HttpMethod.Get, headers);
    }

    public async Task<JsonDocument> GetUserInfoAsync()
    {
        var url = $"{Site}/_api/web/currentuser";
        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };
        return await RequestAsync(url, HttpMethod.Get, headers);
    }

    public async Task<JsonDocument> GetSiteInfoAsync()
    {
        var url = $"{Site}/_api/web";
        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };
        return await RequestAsync(url, HttpMethod.Get, headers);
    }

    public async Task<JsonDocument> SearchItemsAsync(
        string filtro,
        string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista, $"/items?$filter={Uri.EscapeDataString(filtro)}");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };

        return await RequestAsync(url, HttpMethod.Get, headers);
    }

    // =========================
    // Anexos (usando bytes)
    // =========================

    public async Task<JsonDocument> AddAttachmentAsync(
        int itemId,
        string fileName,
        byte[] fileContent,
        bool overwrite = false,
        string lista = null)
    {
        lista ??= ListaAtual;

        // Checa se já existe
        var checkUrl = BuildListUrl(lista, $"/items({itemId})/AttachmentFiles('{fileName}')");
        var checkHeaders = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"}
        };

        var checkRequest = new HttpRequestMessage(HttpMethod.Get, checkUrl);
        foreach (var kvp in checkHeaders)
            checkRequest.Headers.TryAddWithoutValidation(kvp.Key, kvp.Value);

        var checkResp = await _httpClient.SendAsync(checkRequest);

        if (checkResp.IsSuccessStatusCode)
        {
            if (!overwrite)
            {
                // Já existe e não queremos sobrescrever
                return null;
            }
            else
            {
                await RemoveAttachmentAsync(itemId, fileName, lista);
            }
        }

        var url = BuildListUrl(lista, $"/items({itemId})/AttachmentFiles/add(FileName='{fileName}')");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"},
            {"Content-Type", "application/octet-stream"}
        };

        var content = new ByteArrayContent(fileContent);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

        return await RequestAsync(url, HttpMethod.Post, headers, content);
    }

    public async Task<bool> RemoveAttachmentAsync(
        int itemId,
        string fileName,
        string lista = null)
    {
        lista ??= ListaAtual;
        var url = BuildListUrl(lista, $"/items({itemId})/AttachmentFiles('{fileName}')");

        var headers = new Dictionary<string, string>
        {
            {"IF-MATCH", "*"}
        };

        await RequestAsync(url, HttpMethod.Delete, headers);
        return true;
    }

    public async Task<JsonDocument> AddAttachmentToFolderAsync(
        string fileName,
        byte[] fileContent,
        string folderUrl)
    {
        var decodedFolder = Uri.UnescapeDataString(folderUrl);
        var url = BuildSharePointUrl(
            $"GetFolderByServerRelativeUrl('{decodedFolder}')/Files/add(url='{fileName}',overwrite=true)");

        var headers = new Dictionary<string, string>
        {
            {"Accept", "application/json;odata=verbose"},
            {"Content-Type", "application/octet-stream"}
        };

        var content = new ByteArrayContent(fileContent);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

        return await RequestAsync(url, HttpMethod.Post, headers, content);
    }
}
