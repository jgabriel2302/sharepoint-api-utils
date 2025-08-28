import requests
from typing import Optional, Dict, Any

class SPRestApi:
    """
    Classe para interagir com a API REST do SharePoint usando autenticação OAuth2.
    """

    def __init__(self, site_url: str, access_token: str):
        """
        Inicializa a instância da API com a URL do site e o token de acesso OAuth2.
        """
        self.site_url = site_url.rstrip('/')
        self.access_token = access_token
        self.headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
        }

    def _get_list_url(self, list_name: str, endpoint: str = "") -> str:
        """
        Constrói a URL da API para uma lista específica.
        """
        return f"{self.site_url}/_api/web/lists/getbytitle('{list_name}'){endpoint}"

    def get_items(self, list_name: str, params: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """
        Obtém os itens da lista especificada.
        """
        url = self._get_list_url(list_name, "/items")
        if params:
            query_string = '&'.join([f"${key}={value}" for key, value in params.items()])
            url += f"?{query_string}"
        response = requests.get(url, headers=self.headers)
        return response.json()

    def add_item(self, list_name: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Adiciona um item à lista especificada.
        """
        entity_type = f"SP.Data.{list_name.replace(' ', '_x0020_').replace('_', '_x005f_')}ListItem"
        payload = {
            "__metadata": {"type": entity_type},
            **data
        }
        url = self._get_list_url(list_name, "/items")
        response = requests.post(url, headers=self.headers, json=payload)
        return response.json()

    def update_item(self, list_name: str, item_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Atualiza um item existente na lista.
        """
        entity_type = f"SP.Data.{list_name.replace(' ', '_x0020_').replace('_', '_x005f_')}ListItem"
        payload = {
            "__metadata": {"type": entity_type},
            **data
        }
        url = self._get_list_url(list_name, f"/items({item_id})")
        headers = self.headers.copy()
        headers.update({
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        })
        response = requests.post(url, headers=headers, json=payload)
        return response.json()

    def delete_item(self, list_name: str, item_id: int) -> bool:
        """
        Exclui um item da lista.
        """
        url = self._get_list_url(list_name, f"/items({item_id})")
        headers = self.headers.copy()
        headers.update({
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        })
        response = requests.post(url, headers=headers)
        return response.status_code == 204

    def get_user_info(self) -> Dict[str, Any]:
        """
        Obtém informações do usuário atual.
        """
        url = f"{self.site_url}/_api/web/currentuser"
        response = requests.get(url, headers=self.headers)
        return response.json()
    
    def get_access_token(self):
        """
        Obtém o token de acesso OAuth2 para SharePoint usando client_id, client_secret e tenant_id.
        """
        token_url = f"https://accounts.accesscontrol.windows.net/{self.tenant_id}/tokens/OAuth/2"
        resource = f"{self.site}"
        payload = {
            'grant_type': 'client_credentials',
            'client_id': f"{self.client_id}@{self.tenant_id}",
            'client_secret': self.client_secret,
            'resource': f"{resource}@{self.tenant_id}"
        }
    
        response = requests.post(token_url, data=payload)
        if response.status_code == 200:
            token_data = response.json()
            self.access_token = token_data.get('access_token')
            return self.access_token
        else:
            raise Exception(f"Erro ao obter token: {response.status_code} - {response.text}")
