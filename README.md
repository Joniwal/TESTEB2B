# TESTEB2B

Aplicativo Windows em Python + HTML/CSS/JavaScript para consultar, incluir, editar e excluir registros de uma lista corporativa no SharePoint Online, usando Microsoft Graph com a conta do colaborador. Interface em janela nativa (pywebview/WebView2). Não exige servidor compartilhado.

## 1. Prepare a lista na empresa

Crie a lista **TESTEB2B** no site SharePoint desejado:

| Coluna exibida | Tipo | Configuração |
|---|---|---|
| IDCLIENTE | Uma linha de texto | Obrigatório; preserva zeros à esquerda |
| NOME | Uma linha de texto | Obrigatório |
| DATA | Data e hora | Somente data; obrigatório |

Torne a coluna padrão **Título/Title opcional**. Não adicione outros campos obrigatórios sem valor padrão. O ID interno do SharePoint identifica os registros para edição/exclusão; IDCLIENTE é um campo de negócio, não o ID interno. Se precisar impedir clientes duplicados, configure valores exclusivos em IDCLIENTE no SharePoint.

O programa descobre o ID do site, o ID da lista e os nomes internos das três colunas. Não precisa copiar GUIDs da lista. A busca funciona sobre os registros carregados; a leitura percorre todas as páginas do Graph.

O link de Lists pessoal (`lists.live.com`) não é o destino desta versão: crie a lista no SharePoint corporativo.

## 2. Configuração única com a TI

No Microsoft Entra ID da empresa:

1. Registre **TESTEB2B** como aplicativo de locatário único (contas somente da organização).
2. Em Autenticação, adicione plataforma **Aplicativos móveis e de área de trabalho** e URI personalizado `ms-appx-web://microsoft.aad.brokerplugin/SEU_CLIENT_ID`.
3. Adicione permissão **delegada** Microsoft Graph `Sites.ReadWrite.All`. A TI deverá avaliar e conceder o consentimento conforme a política corporativa. Essa permissão é ampla para sites aos quais o usuário tem acesso; o endereço configurado limita o comportamento do programa, não constitui uma restrição de segurança da permissão. Para distribuição além do teste, a TI pode avaliar permissões Selected, que exigem concessões específicas e adaptação dos escopos.
4. Forneça **Directory (tenant) ID** e **Application (client) ID**. Não crie client secret para distribuir com o executável.
5. Garanta que cada colaborador tenha permissão de edição na lista e esteja autorizado a usar o aplicativo, caso o registro exija atribuição de usuários.

O WAM pode aproveitar a conta corporativa conectada no Windows. O login no Windows não substitui o registro/consentimento: pode haver seleção de conta, MFA e regras de acesso condicional. O programa tenta renovação silenciosa durante a sessão; quando não for possível, pede reconexão. Tokens não são escritos pelo aplicativo no .env nem em arquivos; o Windows administra o broker. O botão sair encerra a sessão do programa, sem desconectar a conta do Windows.

## 3. Endereço final no .env

Copie `.env.example` para **`.env`**, na mesma pasta de `TESTEB2B.exe` (ou `app.py` durante desenvolvimento):

```dotenv
TENANT_ID=00000000-0000-0000-0000-000000000000
CLIENT_ID=00000000-0000-0000-0000-000000000000
SHAREPOINT_SITE_URL=https://empresa.sharepoint.com/sites/Equipe
SHAREPOINT_LIST_NAME=TESTEB2B
```

Use o endereço do **site**, sem `/Lists/TESTEB2B/AllItems.aspx`. Pode ser também `/teams/Equipe` ou a raiz do SharePoint. O nome exibido da lista vai no campo separado. O endereço da empresa permanece somente no ambiente corporativo; `.env` é ignorado pelo Git.

Os valores também podem ser digitados na tela. Após uma conexão válida são salvos em `%LOCALAPPDATA%\TESTEB2B\config.json`. Na próxima abertura, valores não vazios do `.env` têm preferência. A configuração salva não contém senha nem token.

## 4. Executar e gerar .exe

Windows 10/11 x64, Python 3.14 x64 para compilar e Microsoft Edge **WebView2 Runtime** instalado. Nas máquinas dos colaboradores não é necessário instalar Python. A TI pode distribuir o WebView2 conforme a política da empresa.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

Para gerar o executável:

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

Resultado: `dist\TESTEB2B.exe`. Distribua junto com o `.env` preenchido **internamente**. O executável não incorpora o .env, então o mesmo binário pode ser levado para a empresa e configurado lá. O build não assina o executável; eventual assinatura/distribuição é responsabilidade da TI.

## 5. Testes

```powershell
python -m unittest discover -s tests -v
```

Os testes simulam o Graph e verificam validação, paginação, mapeamento de colunas, gravação e conflito de edição. Não usam dados da empresa. Para validar na empresa: conectar, criar um cliente com ID `0001`, atualizar a tabela, editar nome/data, conferir no SharePoint e excluir o registro. Teste também com dois usuários editando o mesmo registro: o segundo deve receber aviso de conflito e atualizar a tabela.

Não há repetição automática de gravações após falha de rede: atualize a tabela para conferir se a operação foi efetivada antes de tentar novamente.

## Referências

- [MSAL Python e WAM](https://learn.microsoft.com/en-us/entra/msal/python/advanced/wam)
- [Criar itens pelo Graph e permissões delegadas](https://learn.microsoft.com/en-us/graph/api/listitem-create?view=graph-rest-1.0)
- [Atualizar campos e controle por eTag](https://learn.microsoft.com/en-us/graph/api/listitem-update?view=graph-rest-1.0)
- [Permissões Selected](https://learn.microsoft.com/en-us/graph/permissions-selected-overview)
