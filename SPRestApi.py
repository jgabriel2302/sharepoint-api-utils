import base64
import json
from urllib.parse import quote
import requests

class SPRestApi:
    class Utils:
        @staticmethod
        def base64_para_bytes(base64_str: str, tipo: str = "image/jpeg") -> bytes:
            parts = base64_str.split(",")
            dados = parts[1] if len(parts) > 1 else parts[0]
            return base64.b64decode(dados)

        @staticmethod
        def base64_para_arraybuffer(base64_str: str) -> bytes:
            parts = base64_str.split(",")
            dados = parts[1] if len(parts) > 1 else parts[0]
            return base64.b64decode(dados)

    def __init__(
        self,
        site: str = "https://<seu contoso>.sharepoint.com/sites/<seu site>",
        lista: str | None = None,
        auth_headers: dict | None = None,
    ):
        """
        auth_headers: dicionário com headers de autenticação,
        ex.: {"Authorization": "Bearer <token>"}
        """
        self.site = site.rstrip("/")
        self.lista_atual = lista
        self.type = self.encode_entity_type(self.lista_atual)
        self.session = requests.Session()

        if auth_headers:
            self.session.headers.update(auth_headers)

    # =========================
    # Helpers de URL e tipo
    # =========================

    def set_lista(self, lista_name: str):
        self.lista_atual = lista_name

    def set_list_type(self, list_type: str):
        self.type = list_type
        return self

    def get_lista(self, lista_name: str) -> "SPRestApi":
        return SPRestApi(self.site, lista_name, self.session.headers)

    def encode_entity_type(self, lista: str | None, use_default_type: bool = False) -> str:
        if use_default_type and getattr(self, "type", None):
            return self.type
        lista = "" if lista is None else str(lista)
        encoded = lista.replace(" ", "_x0020_").replace("_", "_x005f_")
        return f"SP.Data.{encoded}ListItem"

    def build_list_url(self, lista: str | None, endpoint: str = "") -> str:
        if not lista:
            raise ValueError("Lista não definida.")
        return f"{self.site}/_api/web/lists/getbytitle('{lista}'){endpoint}"

    def build_sharepoint_url(self, endpoint: str = "") -> str:
        return f"{self.site}/_api/web/{endpoint}"

    # =========================
    # Request genérico
    # =========================

    def request(self, url: str, method: str = "GET", headers: dict | None = None, body=None):
        method = method.upper()
        final_headers = {}
        final_headers.update(self.session.headers)
        if headers:
            final_headers.update(headers)

        if isinstance(body, (dict, list)):
            data = json.dumps(body)
        else:
            data = body

        resp = self.session.request(method, url, headers=final_headers, data=data)
        resp.raise_for_status()
        if resp.content:
            return resp.json()
        return None

    # =========================
    # CRUD de itens
    # =========================

    def add_item(self, data: dict | None = None, lista: str | None = None):
        lista = lista or self.lista_atual
        url = self.build_list_url(lista, "/items")

        headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        }

        payload = {
            "__metadata": {
                "type": self.encode_entity_type(lista, lista == self.lista_atual)
            },
            "Title": "",
        }
        if data:
            payload.update(data)

        return self.request(url, "POST", headers=headers, body=payload)

    def insert_item(self, data: dict | None = None, lista: str | None = None):
        return self.add_item(data, lista)

    def update_item(self, item_id: int, data: dict | None = None, lista: str | None = None):
        lista = lista or self.lista_atual
        url = self.build_list_url(lista, f"/items({item_id})")

        headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
        }

        payload = {
            "__metadata": {
                "type": self.encode_entity_type(lista, lista == self.lista_atual)
            }
        }
        if data:
            payload.update(data)

        # MERGE = POST + X-HTTP-Method
        return self.request(url, "POST", headers=headers, body=payload)

    def delete_item(self, item_id: int, lista: str | None = None) -> bool:
        lista = lista or self.lista_atual
        url = self.build_list_url(lista, f"/items({item_id})")

        headers = {
            "Accept": "application/json;odata=verbose",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
        }

        self.request(url, "POST", headers=headers)
        return True

    def upsert_item(self, data: dict, lista: str | None = None):
        lista = lista or self.lista_atual
        if "Id" in data and data["Id"]:
            item_id = data["Id"]
            new_data = {k: v for k, v in data.items() if k != "Id"}
            return self.update_item(item_id, new_data, lista)
        return self.add_item(data, lista)

    # =========================
    # Leitura de itens e metadados
    # =========================

    def get_items(self, params: dict | None = None, lista: str | None = None):
        lista = lista or self.lista_atual
        base_url = self.build_list_url(lista, "/items")

        if params:
            query_parts = []
            for k, v in params.items():
                query_parts.append(f"${k}={quote(str(v))}")
            query = "&".join(query_parts)
            url = f"{base_url}?{query}"
        else:
            url = base_url

        headers = {"Accept": "application/json;odata=verbose"}
        return self.request(url, "GET", headers=headers)

    def get_item_by_id(self, item_id: int, lista: str | None = None):
        lista = lista or self.lista_atual
        url = self.build_list_url(lista, f"/items({item_id})")

        headers = {"Accept": "application/json;odata=verbose"}
        data = self.request(url, "GET", headers=headers)
        return data.get("d") if data else None

    def get_list_metadata(self, lista: str | None = None):
        lista = lista or self.lista_atual
        url = self.build_list_url(lista, "/fields?$select=Id,EntityPropertyName,Choices,Title,TypeAsString")

        headers = {"Accept": "application/json;odata=verbose"}
        data = self.request(url, "GET", headers=headers)
        return data.get("d", {}).get("results", []) if data else []

    def get_field_metadata_by_name(self, field_name: str, lista: str | None = None):
        lista = lista or self.lista_atual
        url = self.build_list_url(
            lista,
            f"/fields/getbytitle('{field_name}')?$select=Id,EntityPropertyName,Choices,Title,TypeAsString",
        )

        headers = {"Accept": "application/json;odata=verbose"}
        data = self.request(url, "GET", headers=headers)
        return data.get("d") if data else None

    # =========================
    # User / site info
    # =========================

    def get_user_info(self):
        url = f"{self.site}/_api/web/currentuser"
        headers = {"Accept": "application/json;odata=verbose"}
        data = self.request(url, "GET", headers=headers)
        return data.get("d") if data else None

    def get_site_info(self):
        url = f"{self.site}/_api/web"
        headers = {"Accept": "application/json;odata=verbose"}
        data = self.request(url, "GET", headers=headers)
        return data.get("d") if data else None

    # =========================
    # Search OData
    # =========================

    def search_items(self, filtro: str, lista: str | None = None):
        lista = lista or self.lista_atual
        url = self.build_list_url(lista, f"/items?$filter={quote(filtro)}")

        headers = {"Accept": "application/json;odata=verbose"}
        data = self.request(url, "GET", headers=headers)
        return data.get("d", {}).get("results", []) if data else []

    # =========================
    # Anexos
    # =========================

    def add_attachment(
        self,
        item_id: int,
        file_name: str,
        file_content: bytes,
        overwrite: bool = False,
        lista: str | None = None,
    ):
        lista = lista or self.lista_atual

        # Checa se existe
        check_url = self.build_list_url(lista, f"/items({item_id})/AttachmentFiles('{file_name}')")
        check_headers = {"Accept": "application/json;odata=verbose"}

        check_resp = self.session.get(check_url, headers=check_headers)
        if check_resp.ok:
            if not overwrite:
                # Já existe e não quer sobrescrever
                return None
            else:
                self.remove_attachment(item_id, file_name, lista)

        url = self.build_list_url(lista, f"/items({item_id})/AttachmentFiles/add(FileName='{file_name}')")

        headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/octet-stream",
        }

        resp = self.session.post(url, headers=headers, data=file_content)
        resp.raise_for_status()

        return resp.json().get("d")

    def remove_attachment(self, item_id: int, file_name: str, lista: str | None = None) -> bool:
        lista = lista or self.lista_atual
        url = self.build_list_url(lista, f"/items({item_id})/AttachmentFiles('{file_name}')")

        headers = {
            "IF-MATCH": "*",
        }

        resp = self.session.delete(url, headers=headers)
        resp.raise_for_status()
        return True

    def add_attachment_to_folder(self, file_name: str, file_content: bytes, folder_url: str):
        decoded_folder = requests.utils.unquote(folder_url)
        url = self.build_sharepoint_url(
            f"GetFolderByServerRelativeUrl('{decoded_folder}')/Files/add(url='{file_name}',overwrite=true)"
        )

        headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/octet-stream",
        }

        resp = self.session.post(url, headers=headers, data=file_content)
        resp.raise_for_status()
        return resp.json().get("d")
