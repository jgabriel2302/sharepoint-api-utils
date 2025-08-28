# SPRestApi

A classe `SPRestApi` é uma abstração em JavaScript para facilitar a comunicação com a API REST do SharePoint. Ela permite realizar operações como adicionar, atualizar, excluir e consultar itens de listas, além de obter informações do site e do usuário atual.

## 📌 Finalidade

Automatizar e simplificar o uso da API REST do SharePoint em aplicações JavaScript, com suporte para múltiplas listas e reutilização de instância.

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

## 📦 Instalação

Não requer instalação de pacotes externos. Basta importar a classe em seu projeto JavaScript.

## 💡 Exemplos de uso

### JavaScript puro

```javascript
const api = new SPRestApi('https://consoso.sharepoint.com/sites/meusite');
api.setLista('Demandas');

api.addItem({ Title: 'Nova demanda' })
   .then(response => console.log(response));
```

```javascript
const api = new SPRestApi('https://consoso.sharepoint.com/sites/meusite');
const Demandas = api.getLista('Demandas');
const Pessoas = api.getLista('Pessoas');

Demandas.addItem({ Title: 'Nova demanda' })
   .then(response => console.log(response));
Pessoas.removeItem(13)
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

## 🛠️ Requisitos

- SharePoint Online com acesso à API REST.
- Permissões adequadas para leitura/escrita nas listas.
