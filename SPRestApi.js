class SPRestApi {
    static Utils = {
        base64ParaBlob(base64, tipo = 'image/jpeg') {
            const partes = base64.split(',');
            const dados = partes[1];
            const binario = atob(dados);
            const tamanho = binario.length;
            const bytes = new Uint8Array(tamanho);

            for (let i = 0; i < tamanho; i++) {
                bytes[i] = binario.charCodeAt(i);
            }

            return new Blob([bytes], {
                type: tipo
            });
        },
        base64ParaArrayBuffer(base64) {
            const dados = base64.split(',')[1];
            const binario = atob(dados);
            const tamanho = binario.length;
            const buffer = new ArrayBuffer(tamanho);
            const bytes = new Uint8Array(buffer);

            for (let i = 0; i < tamanho; i++) {
                bytes[i] = binario.charCodeAt(i);
            }

            return buffer;
        }
    }
    /**
     * Cria uma instância da API REST do SharePoint.
     * @param {string} site - URL do site SharePoint. https://<seu contoso>.sharepoint.com/sites/<seu site>
     * @param {string|null} lista - Nome da lista padrão (opcional).
     */
    constructor(site = 'https://<seu contoso>.sharepoint.com/sites/<seu site>', lista = null) {
        this.site = site;
        this.listaAtual = lista;
        this.type = this.encodeEntityType(this.listaAtual);
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
     * @param {string} listType - Nome da lista.
     * @returns {SPRestApi}
     */
    setListType(listType) {
        this.type = listType;
        return this;
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
     * @param {boolean} useDefaultType - Utilizar type padrão já definido.
     * @returns {string} Tipo de entidade codificado.
     */
    encodeEntityType(lista, useDefaultType = false) {
        return useDefaultType ? this.type : "SP.Data." + String(lista ?? '').replace(/ /g, '_x0020_').replace(/_/g, '_x005f_') + "ListItem";
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
     * Constrói a URL da API para o sharepoint.
     * @param {string} endpoint - Caminho adicional da API.
     * @returns {string} URL completa da API.
     */
    buildSharePointUrl(endpoint = '') {
        return `${this.site}/_api/web/${endpoint}`;
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
        const response = await fetch(url, {
            method,
            headers,
            body
        });
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
            "__metadata": {
                "type": this.encodeEntityType(lista, lista === this.listaAtual)
            },
            "Title": "",
            ...data
        });
        const response = await this.request(url, "POST", headers, payload);
        return response.error ? false : response;
    }

    /**
     * Adiciona um item à lista.
     * @param {Object} data - Dados do item.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Dados do item criado ou false.
     */
    async insertItem(data = {}, lista = this.listaAtual) {
        return this.addItem(data, lista);
    }


    /**
     * Obtém o ID do usuário no SharePoint a partir do e-mail.
     * @param {string} email - E-mail do usuário.
     * @returns {Promise<number|null>} Retorna o ID do usuário ou null se não encontrado.
     */
    async getUserIdByEmail(email) {
        try {
            const formDigest = await this.getFormDigestValue();
            const url = `${this.site}/_api/web/ensureuser`;

            // Formato Claims para usuários do Azure AD
            const logonName = `i:0#.f|membership|${email}`;

            const response = await fetch(url, {
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "X-RequestDigest": formDigest
                },
                body: JSON.stringify({
                    logonName
                })
            });

            if (!response.ok) {
                console.error("Erro ao buscar usuário:", await response.text());
                return null;
            }

            const data = await response.json();
            return data.d.Id; // ID do usuário no User Information List
        } catch (error) {
            console.error("Erro inesperado em getUserIdByEmail:", error);
            return null;
        }
    }


    /**
     * Adiciona um anexo a um item da lista, com opção de sobrescrever.
     * @param {number} itemId - ID do item.
     * @param {string} fileName - Nome do arquivo.
     * @param {Blob|ArrayBuffer} fileContent - Conteúdo do arquivo.
     * @param {boolean} [overwrite=false] - Se true, sobrescreve o arquivo existente.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Dados do anexo ou false.
     */
    async addAttachment(itemId, fileName, fileContent, overwrite = false, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();

        const checkUrl = this.buildListUrl(lista, `/items(${itemId})/AttachmentFiles('${fileName}')`);
        const checkResponse = await fetch(checkUrl, {
            method: "GET",
            headers: {
                "Accept": "application/json;odata=verbose"
            }
        });

        if (checkResponse.ok) {
            if (overwrite) {
                const removed = await this.removeAttachment(itemId, fileName, lista);
                if (!removed) {
                    console.error("Não foi possível sobrescrever o anexo existente.");
                    return false;
                }
            } else {
                console.warn(`Já existe um anexo com o nome '${fileName}'. Use overwrite=true para substituir.`);
                return false;
            }
        }

        const url = this.buildListUrl(lista, `/items(${itemId})/AttachmentFiles/add(FileName='${fileName}')`);
        const headers = {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": formDigest,
            "Content-Type": "application/octet-stream"
        };

        const response = await fetch(url, {
            method: "POST",
            headers,
            body: fileContent
        });
        if (!response.ok) {
            const error = await response.json();
            console.error("Erro ao adicionar anexo:", error);
            return false;
        }

        const result = await response.json();
        return result.d;
    }

    /**
     * Remove um anexo de um item da lista.
     * @param {number} itemId - ID do item.
     * @param {string} fileName - Nome do arquivo.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<boolean>} True se removido, false caso contrário.
     */
    async removeAttachment(itemId, fileName, lista = this.listaAtual) {
        const formDigest = await this.getFormDigestValue();
        const url = this.buildListUrl(lista, `/items(${itemId})/AttachmentFiles('${fileName}')`);

        const response = await fetch(url, {
            method: "DELETE",
            headers: {
                "X-RequestDigest": formDigest,
                "IF-MATCH": "*"
            }
        });

        if (!response.ok) {
            console.error(`Erro ao remover anexo '${fileName}':`, await response.text());
            return false;
        }
        return true;
    }



    /**
     * Salva um arquivo em uma pasta do SharePoint.
     * @param {string} fileName - Nome do arquivo (ex.: "imagem_1024.jpg").
     * @param {Blob|ArrayBuffer} fileContent - Conteúdo do arquivo.
     * @param {string} folderUrl - Caminho relativo da pasta no SharePoint (ex.: "%2Fsites%2Fcontroladorialongos%2FSiteAssets%2FLists%2F63709c1e-52e2-41f8-8bdc-dffd20fa2ca6").
     * @returns {Promise<Object|boolean>} Dados do arquivo ou false.
     */
    async addAttachmentToFolder(fileName, fileContent, folderUrl) {
        try {
            const formDigest = await this.getFormDigestValue();
            const url = this.buildSharePointUrl(`GetFolderByServerRelativeUrl('${decodeURIComponent(folderUrl)}')/Files/add(url='${fileName}',overwrite=true)`)

            const headers = {
                "Accept": "application/json;odata=verbose",
                "X-RequestDigest": formDigest,
                "Content-Type": "application/octet-stream"
            };

            const response = await fetch(url, {
                method: "POST",
                headers,
                body: fileContent
            });

            if (!response.ok) {
                const error = await response.json();
                console.error("Erro ao salvar arquivo na pasta:", error);
                return false;
            }

            const result = await response.json();
            return result.d;
        } catch (err) {
            console.error("Erro inesperado:", err);
            return false;
        }
    }


    /**
     * Adiciona um item à lista ou atualiza um item existente na lista.
     * @param {Object} data - Dados atualizados.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object|boolean>} Confirmação ou false.
     */
    async upsertItem(data = {}, lista = this.listaAtual) {
        if (data.Id) return this.updateItem(data.Id, Object.keys(data).filter(k => k !== 'Id').reduce((obj, k) => ({
            ...obj,
            [k]: data[k]
        }), {}), lista);
        return this.addItem(data, lista);
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
            "__metadata": {
                "type": this.encodeEntityType(lista, lista === this.listaAtual)
            },
            ...data
        });
        const response = await fetch(url, {
            method: "POST",
            headers,
            body: payload
        });
        if (response.ok) return {
            d: {
                Id: parseInt(id)
            }
        };
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
        const response = await fetch(url, {
            method: "POST",
            headers
        });
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
        return await this.request(url, method, {
            ...defaultHeaders,
            ...headers
        }, body);
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
        const headers = {
            "accept": "application/json;odata=verbose"
        };
        const response = await fetch(url.toString(), {
            method: 'GET',
            headers
        });
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
        const headers = {
            "accept": "application/json;odata=verbose"
        };
        const response = await fetch(url, {
            method: 'GET',
            headers
        });
        const json = await response.json();
        return json.d;
    }

    /**
     * Obtém metadados da lista, como campos e tipos.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object[]>} Metadados da lista.
     */
    async getListMetadata(lista = this.listaAtual) {
        const url = this.buildListUrl(lista, "/fields?$select=Id,EntityPropertyName,Choices,Title,TypeAsString");
        const headers = {
            "accept": "application/json;odata=verbose"
        };
        const response = await fetch(url, {
            method: 'GET',
            headers
        });
        const json = await response.json();
        return json.d.results;
    }

    /**
     * Obtém metadados da lista, como campos e tipos.
     * @param {string} [lista=this.listaAtual] - Nome da lista.
     * @returns {Promise<Object[]>} Metadados da lista.
     */
    async getFieldMetadataByName(fieldName = '', lista = this.listaAtual) {
        const url = this.buildListUrl(lista, "/fields/getbytitle('" + fieldName + "')?$select=Id,EntityPropertyName,Choices,Title,TypeAsString");
        const headers = {
            "accept": "application/json;odata=verbose"
        };
        const response = await fetch(url, {
            method: 'GET',
            headers
        });
        const json = await response.json();
        return json.d;
    }

    /**
     * Obtém informações do usuário atual logado.
     * @returns {Promise<Object>} Dados do usuário.
     */
    async getUserInfo() {
        const url = `${this.site}/_api/web/currentuser`;
        const headers = {
            "accept": "application/json;odata=verbose"
        };
        const response = await fetch(url, {
            method: 'GET',
            headers
        });
        const json = await response.json();
        return json.d;
    }

    /**
     * Obtém informações gerais do site atual.
     * @returns {Promise<Object>} Dados do site.
     */
    async getSiteInfo() {
        const url = `${this.site}/_api/web`;
        const headers = {
            "accept": "application/json;odata=verbose"
        };
        const response = await fetch(url, {
            method: 'GET',
            headers
        });
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
        const headers = {
            "accept": "application/json;odata=verbose"
        };
        const response = await fetch(url, {
            method: 'GET',
            headers
        });
        const json = await response.json();
        return json.d.results;
    }
}

