# Duarte Relats — pacote integrado

Aplicação web de conferência operacional para importação de e-mails exportados e matrizes, com geração de Excel/PDF. A interface começa com uma tela de acesso Microsoft quando o serviço é configurado para autenticação corporativa. Nenhum dado de exemplo é carregado.

## O que este pacote entrega

- Painel sem indicadores fictícios; leitura de `.eml`, `.csv`, `.xlsx` e `.xlsm` de e-mails/metadados.
- Leitura de matriz `.xlsx`/`.xlsm` sem regravar o original, com contagens e hash do conteúdo.
- Exportação de relatório em `.xlsx` e `.pdf`.
- Login web Microsoft Entra ID, validação do fluxo OIDC/MSAL, sessão opaca em cookie `HttpOnly`, `Secure` em HTTPS, `SameSite=Lax`, expiração e logout.
- Restrição de login a endereços explicitamente permitidos por `DUARTE_ALLOWED_USERS`.
- Integração opcional de análise via Azure AI Foundry; depende de configuração aprovada e habilitação explícita.
- Aba de ligações identificada como futura, sem dados simulados.

## Limites importantes

**Este pacote não deve ser tratado como sistema corporativo completo de produção.** Sessões e cache MSAL ficam em memória do processo: reinício encerra sessões e múltiplas réplicas exigem armazenamento compartilhado seguro. Ainda faltam persistência multiusuário, auditoria central, rate limiting, CSRF robusto, gestão de papéis e análise formal de privacidade/retensão.

O login Microsoft não conecta nem lê o Outlook. Microsoft Graph para listar mensagens, caixas, paginação, anexos e tratamento de limites ainda não está implementado. Para o login, a permissão delegada mínima é `User.Read`; não mantenha `Mail.Read` apenas por causa desta versão. Não há integração Google, Duarte Performance ou telefonia.

A IA é desativada por padrão e só analisa arquivos selecionados. O corpo das mensagens é enviado ao endpoint Azure configurado; obtenha aprovação organizacional antes de usar dados reais. Anexos não são analisados. Não use o gate compartilhado de IA como substituto de autenticação corporativa. Valide toda saída da IA contra as fontes originais.

## Deploy no Railway

1. Coloque **o conteúdo desta pasta** na raiz do repositório conectado ao serviço Railway, ou configure o Root Directory para a pasta do projeto. Confirme que `server.py`, `index.html`, `requirements.txt`, `Procfile` e `.python-version` estão no commit efetivamente implantado.
2. Configure as variáveis protegidas listadas abaixo no serviço correto. Não as inclua no GitHub, no chat nem em capturas.
3. O callback no Entra precisa ser plataforma **Web** e corresponder exatamente a `https://duartelats-production.up.railway.app/auth/callback`. Se o domínio mudar, atualize tanto o Entra quanto `DUARTE_PUBLIC_BASE_URL`.
4. Faça deploy e confirme nos logs a instalação das dependências e a inicialização do serviço. Teste `/api/health`; ele não revela credenciais. Acesse `/` para ver a tela de entrada e então escolha **Entrar com Microsoft**.
5. Se houver erro no retorno, confira os logs de runtime logo após nova tentativa. Não compartilhe URL com `code`/`state`, cookies, tokens ou valores de variáveis.

### Microsoft Entra ID

- Registre um aplicativo para o locatário corporativo e adicione URI de retorno como plataforma Web.
- Configure a permissão **delegada** Microsoft Graph `User.Read` para a identidade do usuário; solicite consentimento conforme a política da organização.
- Acesso é permitido somente aos e-mails normalizados que estiverem na lista `DUARTE_ALLOWED_USERS`.
- Este fluxo não concede acesso de leitura de e-mail e não implementa Graph.

### Variáveis obrigatórias no Railway

| Variável | Valor |
|---|---|
| `DUARTE_ENTRA_TENANT_ID` | ID do locatário no Entra |
| `DUARTE_ENTRA_CLIENT_ID` | ID do aplicativo (cliente) |
| `DUARTE_ENTRA_CLIENT_SECRET` | Valor do segredo de cliente, não seu ID |
| `DUARTE_SESSION_SECRET` | Segredo aleatório criptograficamente forte com pelo menos 32 caracteres |
| `DUARTE_AUTH_REQUIRED` | `true` |
| `DUARTE_PUBLIC_BASE_URL` | URL HTTPS pública do serviço |
| `DUARTE_ALLOWED_USERS` | Lista separada por vírgula de e-mails autorizados, sem espaços excedentes |

Gere o segredo de sessão localmente no computador de TI e cole-o diretamente no cofre de variáveis do serviço. Não reutilize o segredo de cliente Entra como segredo de sessão. O `.env.example` contém apenas placeholders e nunca deve receber credenciais reais em commit.

## Execução local

Requer Python 3.12. Instale dependências e inicie:

```bash
python -m pip install -r requirements.txt
python server.py
```

Acesse `http://127.0.0.1:8000`. Sem variáveis de autenticação, o serviço roda em modo local para teste de arquivos. **Não exponha esse modo local à internet.** Para simular autenticação, não crie valores falsos: configure um aplicativo de teste Entra e uma lista de usuários de teste autorizados.

## Azure AI Foundry (opcional)

A integração é desabilitada por padrão. A TI deve configurar, no ambiente protegido do servidor, `AZURE_OPENAI_ENDPOINT` (HTTPS), `AZURE_OPENAI_DEPLOYMENT`, `AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_API_VERSION`, `DUARTE_AI_ANALYSIS_ENABLED=true` e `DUARTE_AI_ACCESS_TOKEN`. Confirme com o time de plataforma que endpoint, deployment e versão REST são suportados pelo recurso. O segredo compartilhado da análise não é conta Microsoft nem controle multiusuário. Não habilite IA até aprovar finalidade, acesso, tratamento e transferência dos dados.

## Rotas para verificação

- `GET /` — painel local ou página de entrada corporativa.
- `GET /auth/login` — inicia o login Microsoft quando configurado.
- `GET /auth/callback` — retorno OAuth/OIDC.
- `GET /api/auth/status` — estado da sessão, sem tokens.
- `POST /auth/logout` — encerra sessão.
- `GET /api/health` — status básico sem segredos.
- `POST /api/emails`, `/api/matrix`, `/api/export/xlsx`, `/api/export/pdf` — importação e exportação.

## Segurança e dados

- Não inclua credenciais no código ou arquivos versionados. Revogue imediatamente credenciais expostas.
- Use apenas HTTPS publicamente e configure a URI de retorno exatamente.
- Não carregue dados reais até validar autenticação, lista de usuários, permissões, controles de acesso e obrigações LGPD da organização.
- Os arquivos são processados na memória da requisição e não há persistência operacional intencional. O cache OAuth e sessões persistem apenas na memória do processo.
- Arquivos de origem devem ser revisados e autorizados; não trate fórmulas externas sem valor calculado como zero ou como conclusão.
- Datas sem fuso em CSV/XLSX não são convertidas automaticamente. A seleção de fuso na interface é informativa para essas fontes.

## Verificação local

```bash
python -m py_compile server.py
python -m unittest discover -s tests -v
```

A validação local não substitui um teste de login real no Entra nem confirma o deploy no Railway. O login deve ser testado após configurar variáveis protegidas e o callback correto.

## Não incluído

Conexão Graph e leitura de e-mails, análise de anexos/OCR, persistência PostgreSQL, histórico multiusuário, reconciliação real com Duarte Performance, OAuth Google, conexão de telefonia e regras oficiais do semáforo. Ausência de informação deve permanecer “não informado”/“a confirmar”, não zero presumido.
