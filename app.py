from __future__ import annotations

import ctypes
import json
import os
from pathlib import Path
import sys
import threading
import uuid
from datetime import date
from urllib.parse import urlparse, unquote, quote

import requests

GRAPH = 'https://graph.microsoft.com/v1.0'
SCOPES = ['Sites.ReadWrite.All']
CONFIG = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'TESTEB2B' / 'config.json'


def validate_config(value):
    result = {k: str(value.get(k, '')).strip() for k in
              ('tenant_id', 'client_id', 'site_url', 'list_name')}
    for key in ('tenant_id', 'client_id'):
        try:
            uuid.UUID(result[key])
        except ValueError:
            raise ValueError('Informe Tenant ID e Client ID válidos, fornecidos pela TI.')
    parsed = urlparse(result['site_url'])
    if (parsed.scheme != 'https' or not parsed.hostname or
            not parsed.hostname.endswith('.sharepoint.com') or parsed.username or
            parsed.password or parsed.port not in (None, 443)):
        raise ValueError('Use o endereço HTTPS do site da empresa em sharepoint.com.')
    path = unquote(parsed.path).rstrip('/')
    if '/lists/' in path.lower() or path.lower().endswith('.aspx'):
        raise ValueError('Informe o endereço do SITE, por exemplo https://empresa.sharepoint.com/sites/Equipe, sem /Lists/ ou AllItems.aspx.')
    result['site_url'] = 'https://' + parsed.hostname + path
    if not result['list_name']:
        raise ValueError('Informe o nome da lista.')
    return result


def validate_record(value):
    fields = {k: str(value.get(k, '')).strip() for k in ('IDCLIENTE', 'NOME', 'DATA')}
    if not fields['IDCLIENTE'] or not fields['NOME']:
        raise ValueError('Preencha IDCLIENTE e NOME.')
    if len(fields['IDCLIENTE']) > 255 or len(fields['NOME']) > 255:
        raise ValueError('IDCLIENTE e NOME devem ter até 255 caracteres.')
    try:
        date.fromisoformat(fields['DATA'])
    except ValueError:
        raise ValueError('Informe uma data válida.')
    fields['DATA'] += 'T00:00:00Z'
    return fields


class GraphClient:
    def __init__(self, config, token):
        self.config, self.token = config, token
        self.session = requests.Session()
        self.base = None
        self.columns = {}

    def request(self, method, path, **kwargs):
        url = path if path.startswith('https://') else GRAPH + path
        if not url.startswith(GRAPH + '/'):
            raise ValueError('Endereço de paginação inválido.')
        headers = {'Authorization': 'Bearer ' + self.token()}
        headers.update(kwargs.pop('extra_headers', {}))
        response = self.session.request(method, url,
            headers=headers, timeout=(10, 45), **kwargs)
        if not response.ok:
            messages = {401: 'Sessão expirada. Conecte novamente.',
                        403: 'Acesso negado. Verifique permissões do usuário, consentimento e registro do aplicativo com a TI.',
                        404: 'Site, lista ou registro não encontrado. Confira a configuração e atualize a tabela.',
                        412: 'Este registro foi alterado por outra pessoa. Atualize a tabela antes de tentar novamente.',
                        429: 'Microsoft Graph está limitando as requisições. Aguarde e tente novamente.'}
            raise ValueError(messages.get(response.status_code,
                f'Microsoft Graph retornou erro {response.status_code}. Verifique os tipos das colunas e os campos obrigatórios da lista.'))
        return response.json() if response.content else {}

    def all(self, path):
        values = []
        while path:
            page = self.request('GET', path)
            values.extend(page.get('value', []))
            path = page.get('@odata.nextLink')
        return values

    def connect(self):
        parsed = urlparse(self.config['site_url'])
        path = quote(unquote(parsed.path).rstrip('/'), safe='/')
        site = self.request('GET', '/sites/' + parsed.hostname + (':' + path if path else ''))
        lists = self.all('/sites/' + site['id'] + '/lists?$select=id,displayName')
        matches = [x for x in lists if x['displayName'].casefold() == self.config['list_name'].casefold()]
        if len(matches) != 1:
            raise ValueError('Não foi possível identificar a lista pelo nome neste site. Confira o nome exibido no SharePoint.')
        self.base = '/sites/' + site['id'] + '/lists/' + matches[0]['id']
        columns = self.all(self.base + '/columns')
        for field in ('IDCLIENTE', 'NOME', 'DATA'):
            found = [x for x in columns if x.get('displayName', '').casefold() == field.casefold()]
            if len(found) != 1:
                raise ValueError('Coluna ausente ou ambígua: ' + field)
            column = found[0]
            expected = 'dateTime' if field == 'DATA' else 'text'
            if expected not in column or column.get('readOnly'):
                raise ValueError(f'A coluna {field} deve ser editável e do tipo ' + ('Data e hora.' if field == 'DATA' else 'Uma linha de texto.'))
            if field == 'DATA' and column['dateTime'].get('format') != 'dateOnly':
                raise ValueError('Configure a coluna DATA como somente data, sem horário.')
            self.columns[field] = column['name']
        supported = set(self.columns.values())
        for column in columns:
            if (column.get('required') and not column.get('readOnly') and
                    column['name'] not in supported and not column.get('defaultValue')):
                raise ValueError('Torne opcional o campo obrigatório adicional: ' + column.get('displayName', column['name']))

    def rows(self):
        rows = self.all(self.base + '/items?$expand=fields&$top=200')
        return [{'id': x['id'], 'etag': x.get('eTag'), **{
            k: str(x.get('fields', {}).get(v, '') or '')[:10] if k == 'DATA'
            else str(x.get('fields', {}).get(v, '') or '')
            for k, v in self.columns.items()}} for x in rows]

    def save(self, record, item_id=None, etag=None):
        fields = {self.columns[k]: v for k, v in validate_record(record).items()}
        if item_id:
            if not etag:
                raise ValueError('Atualize a tabela antes de editar este registro.')
            return self._update(item_id, fields, etag)
        return self.request('POST', self.base + '/items', json={'fields': fields})

    def _update(self, item_id, fields, etag):
        return self.request('PATCH', self.base + '/items/' + quote(str(item_id), safe='') + '/fields',
                            json=fields, extra_headers={'If-Match': etag})

    def delete(self, item_id, etag):
        if not etag:
            raise ValueError('Atualize a tabela antes de excluir este registro.')
        return self.request('DELETE', self.base + '/items/' + quote(str(item_id), safe=''),
                            extra_headers={'If-Match': etag})


class Api:
    def __init__(self):
        self.lock = threading.Lock()
        self.client = self.auth = self.account = None
        self.config = {}
        if CONFIG.exists():
            try:
                self.config = json.loads(CONFIG.read_text(encoding='utf-8'))
            except (ValueError, OSError):
                pass
        from dotenv import dotenv_values
        folder = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        env = dotenv_values(folder / '.env')
        for field, key in [('tenant_id', 'TENANT_ID'), ('client_id', 'CLIENT_ID'),
                           ('site_url', 'SHAREPOINT_SITE_URL'), ('list_name', 'SHAREPOINT_LIST_NAME')]:
            if env.get(key):
                self.config[field] = env[key]
        self.config.setdefault('list_name', 'TESTEB2B')

    def _run(self, callback):
        with self.lock:
            try:
                return {'ok': True, 'data': callback()}
            except requests.RequestException:
                return {'ok': False, 'error': 'Falha de conexão. Verifique internet, VPN e proxy. Se estava salvando, atualize a tabela antes de repetir a operação.'}
            except Exception as exc:
                return {'ok': False, 'error': str(exc)}

    def settings(self):
        return {'ok': True, 'data': self.config}

    def connect(self, config):
        def action():
            import msal
            self.client = None
            cfg = validate_config(config)
            auth = msal.PublicClientApplication(cfg['client_id'],
                authority='https://login.microsoftonline.com/' + cfg['tenant_id'],
                enable_broker_on_windows=sys.platform == 'win32')
            kwargs = {}
            if sys.platform == 'win32':
                user32 = ctypes.windll.user32
                user32.GetForegroundWindow.restype = ctypes.c_void_p
                kwargs['parent_window_handle'] = user32.GetForegroundWindow()
            result = auth.acquire_token_interactive(SCOPES, **kwargs)
            if 'access_token' not in result:
                raise ValueError('Login não concluído. Verifique o registro do aplicativo e o consentimento com a TI. Código: ' + result.get('error', 'cancelado'))
            accounts = auth.get_accounts()
            username = result.get('id_token_claims', {}).get('preferred_username', '')
            self.auth = auth
            self.account = next((a for a in accounts if a.get('username') == username), accounts[0] if accounts else None)
            client = GraphClient(cfg, self._token)
            client.connect()
            CONFIG.parent.mkdir(parents=True, exist_ok=True)
            CONFIG.write_text(json.dumps(cfg, indent=2), encoding='utf-8')
            self.config, self.client = cfg, client
            return {'username': username}
        return self._run(action)

    def _token(self):
        result = self.auth.acquire_token_silent(SCOPES, account=self.account) if self.account else None
        if not result or 'access_token' not in result:
            raise ValueError('Conecte novamente para renovar sua sessão.')
        return result['access_token']

    def _connected(self):
        if not self.client:
            raise ValueError('Conecte ao Microsoft 365 primeiro.')
        return self.client

    def rows(self):
        return self._run(lambda: self._connected().rows())

    def save(self, record, item_id=None, etag=None):
        return self._run(lambda: self._connected().save(record, item_id, etag))

    def delete(self, item_id, etag):
        return self._run(lambda: self._connected().delete(item_id, etag))

    def disconnect(self):
        with self.lock:
            self.client = self.auth = self.account = None
        return {'ok': True}


if __name__ == '__main__':
    import webview
    root = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))
    webview.create_window('TESTEB2B • SharePoint', str(root / 'ui' / 'index.html'),
                          js_api=Api(), width=1100, height=760, min_size=(800, 600))
    webview.start(debug=False)
