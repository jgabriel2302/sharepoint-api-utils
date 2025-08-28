# SPRestApi

A classe `SPRestApi` é uma abstração para facilitar a comunicação com a API REST do SharePoint. Ela está disponível em **JavaScript**, **C#** e **Python**, permitindo realizar operações como adicionar, atualizar, excluir e consultar itens de listas, além de obter informações do site e do usuário atual.

---

## 📌 Finalidade

Automatizar e simplificar o uso da API REST do SharePoint em aplicações modernas, com suporte para múltiplas listas, autenticação via OAuth2 (em C# e Python) e reutilização de instância.

---

## 🚀 Funcionalidades

| Método | Descrição |
|--------|-----------|
| `setLista(listaName)` | Define a lista atual. |
| `getLista(listaName)` | Cria uma nova instância com lista definida. |
| `addItem(data, lista)` | Adiciona um item à lista. |
| `updateItem(id, data, lista)` | Atualiza um item existente. |
| `deleteItem(id, lista)` | Exclui um item da lista. |
| `getItems(params, lista)` | Obtém itens da lista. |
| `getItemById(id, lista)` | Obtém um item específico. |
| `addAttachment(itemId, fileName, fileContent, lista)` | Adiciona anexo a um item. |
| `getListMetadata(lista)` | Obtém metadados da lista. |
| `getUserInfo()` | Obtém dados do usuário atual. |
| `getSiteInfo()` | Obtém dados do site atual. |
| `searchItems(filtro, lista)` | Pesquisa itens com filtro OData. |
| `anyRequest(api, method, body, headers)` | Executa requisição arbitrária. |
| `getAccessToken()` *(C# e Python)* | Obtém o token OAuth2 para autenticação. |

---

## 📦 Instalação

### JavaScript
Basta importar a classe em seu projeto.

### C#
Referencie `System.Net.Http`, `Newtonsoft.Json` e use `HttpClient` com OAuth2.

### Python
Instale `requests`:
```bash
pip install requests
```

---

## 💡 Exemplos de uso

### JavaScript puro

```javascript
const api = new SPRestApi('https://consoso.sharepoint.com/sites/meusite');
api.setLista('Demandas');

api.addItem({ Title: 'Nova demanda' })
   .then(response => console.log(response));
```

### Vue.js

```javascript
<script>
import SPRestApi from './SPRestApi.js';

export default {
  data() {
    return {
      api: new SPRestApi('https://consoso.sharepoint.com/sites/meusite')
    };
  },
  mounted() {
    this.api.setLista('Projetos');
    this.api.getItems({ top: 5 }).then(items => {
      console.log('Itens:', items);
    });
  }
};
</script>
```

### React

```javascript
import React, { useEffect } from 'react';
import SPRestApi from './SPRestApi.js';

function App() {
  useEffect(() => {
    const api = new SPRestApi('https://consoso.sharepoint.com/sites/meusite');
    api.setLista('Tarefas');
    api.getItems({ top: 10 }).then(items => {
      console.log('Itens da lista:', items);
    });
  }, []);

  return <div>Veja o console para os dados da lista SharePoint</div>;
}

export default App;
```

---

## 🐍 Python

```python
from SPRestApi import SPRestApi

api = SPRestApi(site='https://consoso.sharepoint.com/sites/meusite',
                client_id='seu-client-id',
                client_secret='seu-client-secret',
                tenant_id='seu-tenant-id')

token = api.get_access_token()
items = api.get_items('Demandas')
print(items)
```

---

## 💻 C#

```csharp
var api = new SPRestApi("https://consoso.sharepoint.com/sites/meusite",
                        "seu-client-id",
                        "seu-client-secret",
                        "seu-tenant-id");

string token = await api.GetAccessTokenAsync();
string itemsJson = await api.GetItemsAsync("Demandas");
Console.WriteLine(itemsJson);
```

---

## 🛠️ Requisitos

- SharePoint Online com acesso à API REST.
- Permissões adequadas para leitura/escrita nas listas.
- Para C# e Python: registro de aplicativo no Azure AD com permissões para SharePoint.

---
