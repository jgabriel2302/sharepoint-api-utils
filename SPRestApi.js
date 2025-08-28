class SPRestApi {
    /**
     * Cria uma instância da API REST do SharePoint.
     * @param {string} site - URL do site SharePoint. https://<seu contoso>.sharepoint.com/sites/<seu site>
     * @param {string|null} lista - Nome da lista padrão (opcional).
     */
    constructor(site = 'https://<seu contoso>.sharepoint.com/sites/<seu site>', lista = null) {
        this.site = site;
        this.listaAtual = lista;
    }

    /**
     * Define a lista atual para operações subsequentes.
     * @param {string} listaName - Nome da lista.
     */
    setLista(listaName) {
        this.listaAtual = listaName;
    }

    /**
     * Cria uma nova instância da API com uma lista específica.
     * @param {string} listaName - Nome da lista.
     * @returns {SPRestApi} Nova instância com lista definida.
     */
    getLista(listaName) {
        return new SPRestApi(this.site, listaName);
    }

    /**
     * Codifica o nome da lista para o formato esperado pelo SharePoint.
     * @param {string} lista - Nome da lista.
     * @returns {string} Tipo de entidade codificado.
     */
    encodeEntityType(lista) {
        return "SP.Data." + lista.replace(/ /g, '_x0020_').replace(/_/g, '_x005f_') + "ListItem";
    }

    /**
     * Constrói a URL da API para uma lista.
     * @param {string} lista - Nome da lista.
     * @param {string} endpoint - Caminho adicional da API.
     * @returns {string} URL completa da API.
     */
    buildListUrl(lista, endpoint = '') {
        if (!lista) throw new Error("Lista não definida.");
        return `${this.site}/_api/web/lists/getbytitle('${lista}')${endpoint}`;
    }

    /**
     * Executa uma requisição HTTP genérica.
     * @param {string} url - URL da requisição.
     * @param {string} method - Método HTTP.
     * @param {Object} headers - Cabeçalhos da requisição.
     * @param {any} body - Corpo da requisição.
     * @returns {Promise<Object>} Resposta da API.
     */
    async request(url, method = 'GET', headers = {}, body = null) {
        const response = await fetch(url, { method, headers, body });
        const json = await response.json();
        return json;
    }

    /**
     * Obtém o valor do Form Digest necessário para requisições POST.
     * @returns {Promise<string>} Valor do Form Digest.
     */
    async getFormDigestValue() {
        try {
            const url = `${this.site}/_api/contextinfo`;
            const headers = {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            };
            const data = await this.request(url, "POST", headers);
            return data.d.GetContextWebInformation.FormDigestValue;
        } catch (error) {
            console.error("Erro ao obter o Form Digest:", error);
            return _spPageContextInfo.formDigestValue;
        }
    }

    /**
     * Adiciona um item à lista.
     * @param {Object} data - Dados do item.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Dados do item criado ou false.
     */
    async addItem(data = {}, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, "/items");
        const headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigest
        };
        const payload = JSON.stringify({
            "__metadata": { "type": this.encodeEntityType(lista) },
            "Title": "",
            ...data
        });
        const response = await this.request(url, "POST", headers, payload);
        return response.error ? false : response;
    }

    /**
     * Adiciona um anexo a um item da lista.
     * @param {number} itemId - ID do item.
     * @param {string} fileName - Nome do arquivo.
     * @param {Blob|ArrayBuffer} fileContent - Conteúdo do arquivo.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Dados do anexo ou false.
     */
    async addAttachment(itemId, fileName, fileContent, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, `/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`);
        const headers = {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": formDigest,
            "Content-Type": "application/octet-stream"
        };
        const response = await fetch(url, { method: "POST", headers, body: fileContent });
        if (!response.ok) {
            const error = await response.json();
            console.error("Erro ao adicionar anexo:", error);
            return false;
        }
        const result = await response.json();
        return result.d;
    }

    /**
     * Atualiza um item existente na lista.
     * @param {number} id - ID do item.
     * @param {Object} data - Dados atualizados.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Confirmação ou false.
     */
    async updateItem(id, data = {}, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, `/items(${id})`);
        const headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": formDigest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        };
        const payload = JSON.stringify({
            "__metadata": { "type": this.encodeEntityType(lista) },
            ...data
        });
        const response = await fetch(url, { method: "POST", headers, body: payload });
        if (response.ok) return { d: { Id: parseInt(id) } };
        const errorData = await response.json();
        console.error('Erro detalhado do SharePoint:', errorData.error.message.value);
        return false;
    }

    /**
     * Exclui um item da lista.
     * @param {number} id - ID do item.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<boolean>} True se sucesso.
     */
    async deleteItem(id, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, `/items(${id})`);
        const headers = {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": formDigest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        };
        const response = await fetch(url, { method: "POST", headers });
        if (!response.ok) throw new Error(await response.text());
        return true;
    }

    /**
     * Executa qualquer requisição arbitrária à API do SharePoint.
     * @param {string} api - Caminho da API.
     * @param {string} [method="GET"] - Método HTTP.
     * @param {any} [body=null] - Corpo da requisição.
     * @param {Object} [headers={}] - Cabeçalhos adicionais.
     * @returns {Promise<Object>} Resposta da API.
     */
    async anyRequest(api, method = "GET", body = null, headers = {}) {
        const url = `${this.site}/_api/${api}`;
        const defaultHeaders = {
            "accept": "application/json;odata=verbose",
            "accept-language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            "charset": "UTF-8"
        };
        return await this.request(url, method, { ...defaultHeaders, ...headers }, body);
    }

    /**
     * Obtém itens da lista com parâmetros opcionais de consulta.
     * @param {Object} [params={}] - Parâmetros de consulta.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object>} Lista de itens.
     */
    async getItems(params = {}, lista = this.listaAtual) {
        const url = new URL(this.buildListUrl(lista, "/items"));
        for (const parameter in params) {
            url.searchParams.append(`$${parameter}`, params[parameter]);
        }
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url.toString(), { method: 'GET', headers });
        return await response.json();
    }

    /**
     * Recupera um item específico da lista pelo ID.
     * @param {number} id - ID do item.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object>} Item recuperado.
     */
    async getItemById(id, lista = this.listaAtual) {
        const url = this.buildListUrl(lista, `/items(${id})`);
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d;
    }

    /**
     * Obtém metadados da lista, como campos e tipos.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object[]>} Metadados da lista.
     */
    async getListMetadata(lista = this.listaAtual) {
        const url = this.buildListUrl(lista, "/fields");
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d.results;
    }

    /**
     * Obtém informações do usuário atual logado.
     * @returns {Promise<Object>} Dados do usuário.
     */
    async getUserInfo() {
        const url = `${this.site}/_api/web/currentuser`;
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d;
    }

    /**
     * Obtém informações gerais do site atual.
     * @returns {Promise<Object>} Dados do site.
     */
    async getSiteInfo() {
        const url = `${this.site}/_api/web`;
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d;
    }

    /**
     * Pesquisa itens na lista com base em um filtro OData.
     * @param {string} filtro - Filtro OData.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object[]>} Itens filtrados.
     */
    async searchItems(filtro, lista = this.listaAtual) {
        const url = this.buildListUrl(lista, `/items?$filter=${encodeURIComponent(filtro)}`);
        const headers = { "accept": "application/json;odata=verbose" };
        const response = await fetch(url, { method: 'GET', headers });
        const json = await response.json();
        return json.d.results;
    }
}