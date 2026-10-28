# DUARTE REPORTS — versão local executável

Aplicação web local para leitura de arquivos selecionados, análise de metadados, importação de matriz e geração de Excel/PDF. Base visual: tons Duarte (#001E57, #0A2540 e #FF9200), estrutura do HTML de referência adaptada para relatórios. O HTML original depende de estilos/recursos não incluídos no anexo; nenhuma identidade visual externa ausente foi presumida.

## O que está realmente implementado

- Dashboard inicia sem números fictícios.
- Importação local explícita de mensagens `.eml` individuais e arquivos `.csv`, `.xlsx` ou `.xlsm` com metadados. Limite: 25 MB por arquivo. `.pst` e `.msg` não são suportados nesta versão.
- Datas inicial/final aplicadas localmente; fuso selecionável/registrado para explicitar a leitura, mas datas sem offset em CSV/XLSX não são convertidas automaticamente.
- Campos de e-mail aceitos por cabeçalhos reconhecidos: data/data-hora, remetente/from, destinatários/to/cc, assunto/subject, direção/direction, cliente/operadora e identificadores de mensagem/conversa. A associação de cliente só usa coluna explícita; não deduz por assunto/domínio/destinatário.
- Deduplicação por Message-ID, ou hash de data, remetente, destinatários e assunto. Cada resposta segue sendo mensagem separada; thread/conversa não é tratada como unidade de mensagem.
- Importação `.xlsx`/`.xlsm` de matriz sem alterar o original, identifica abas e cabeçalhos COD/CLIENTE/OPERADORA e exibe registros, contagens e hash SHA-256. Leitura de fórmulas usa valores salvos no workbook. Se valores dependem de links/fórmulas externas sem cache atualizado, resultados podem faltar; validar no Excel fonte.
- Semáforo provisório por datas em colunas com “contato”; limites ajustáveis apenas na interface desta sessão. Não constitui regra operacional aprovada nem tem histórico persistente.
- Exportação `.xlsx` (abas resumo, detalhamento, consolidado, matriz/revisão e metodologia) e PDF narrativo derivados do conteúdo importado. O sistema não guarda arquivos após processar.
- Aba de ligações futura, sem dados ou métricas fictícias.

## Não configurado / não afirmar como funcional

Login Google OAuth/OIDC, Microsoft Graph, permissão de caixa Outlook, persistência compartilhada, auditoria central, armazenamento seguro, API do Duarte Performance e integração telefônica. O estado de autenticação é “não configurado”; não há login simulado.

A planilha de modelo enviada inclui “Resumo Executivo”, “MedLight — Clínica MedLight (1010)” e “EvCiti — Centro de Infusão (1044)”; os detalhes afirmam “Nenhum e-mail encontrado para este cliente no período” e totais apresentados são zero. São estados do modelo, não dados operacionais atuais. PDF de referência é do Duarte Transcribe e serve apenas como exemplo de organização narrativa.

## Execução local

Requer Python 3.11 e dependências `openpyxl` e `reportlab` (já disponíveis no runtime desta entrega). No terminal:

```bash
cd outputs/duarte-reports
python server.py
```

Abrir `http://127.0.0.1:8000`. O serviço escuta em todas as interfaces da máquina (`0.0.0.0`); para uso estritamente local, limitar o acesso a localhost/firewall da máquina. Não expor este servidor à internet: não há autenticação, CSRF, controle multiusuário ou sessão segura.

### Verificação de dependências num ambiente novo

```bash
python -m pip install openpyxl reportlab
python server.py
```

Não existe `.env` com credenciais, porque nenhuma integração OAuth/Graph está habilitada. Variáveis aceitas: `PORT` para porta local; não são segredos.

## Arquitetura recomendada para produção

- Front-end estático servido por CDN/servidor web com CSP e configuração HTTPS.
- Back-end autenticado para OAuth/OIDC Google, sessões com cookies HttpOnly/Secure/SameSite, autorização por usuário/perfil e auditoria.
- Microsoft Graph via fluxo OAuth com consentimento explícito e escopos mínimos de leitura, tokens protegidos apenas no servidor/secret store, seleção explícita de caixa/pasta e expiração/revogação. Login Google nunca substitui autorização Microsoft.
- PostgreSQL para usuários, importações, vínculo revisado, regras de semáforo, execução/exportação e auditoria; armazenamento de objetos criptografado e de retenção curta se arquivos forem indispensáveis. Não armazenar corpos de e-mails por padrão.
- Deploy com serviço separado de aplicação, banco gerenciado e armazenamento persistente; migrações versionadas, backup automatizado e teste periódico de restauração. Nenhum Railway/PostgreSQL está conectado/configurado neste pacote.
- Duarte Performance por exportação autorizada inicialmente; API exige endpoints, autenticação, permissões e mapeamento de métricas confirmados.

### Esquema mínimo recomendado

`users(id, email, display_name, role, created_at)`; `imports(id, user_id, kind, source_name, sha256, read_at, reference_start, reference_end, row_count, ignored_count, errors_json, version)`; `email_metadata(id, import_id, message_key, sent_at, sender, recipients_json, subject, direction, thread_id, customer_code, customer_label, match_method, reviewed_by, reviewed_at)`; `matrix_records(id, import_id, row_key, customer_code, operator_id, operator_name, source_date, fields_json)`; `rules(id, scope, green_days, orange_days, red_days, effective_at, user_id)`; `reconciliations(id, matrix_import_id, performance_import_id, metric, reports_value, performance_value, delta, status)`; `audit_events(id, user_id, action, object_type, object_id, occurred_at, details_json)`. Retenção, papéis e política de acesso: a confirmar com Duarte Gestão e encarregado LGPD.

## Campos e itens a confirmar com o usuário

1. Definição de “contato válido”, inclusão de e-mails recebidos, critérios de associação por código, domínio e apelidos, regras de mensagens automáticas e de conversas.
2. Prazos/limites verde, laranja, vermelho e segmentação por rotina/cliente; fontes válidas de última data de contato.
3. Cabeçalho/linha de dados autoritativos da matriz e campos de código estável. Planilha apresenta fórmulas externas para cliente, ID, prioridade, código ANS e status, além de datas e campos de contato; resultados podem não estar resolvidos fora do Excel.
4. Mapeamento das métricas/colunas de Duarte Performance, período e método de comparação.
5. Política de retenção, perfis de acesso, provedores/domínios OAuth autorizados e consentimentos administrativos Graph.
6. Fonte telefônica e campos autorizados para fase futura.

## Checklist básico

- [ ] Testar importação de EML exportado com data/remetente/assunto.
- [ ] Testar CSV UTF-8 com cabeçalhos reconhecidos e período de datas.
- [ ] Testar XLSX/XLSM sem senha e comparar linha a linha com origem.
- [ ] Testar workbook com fórmulas sem cache (não tratar valor vazio como zero).
- [ ] Confirmar deduplicação quando Message-ID está ausente/presente.
- [ ] Comparar abas e fórmulas críticas no Excel original; o arquivo fonte nunca deve ser regravado.
- [ ] Conferir Excel/PDF exportados contra a mesma importação.
- [ ] Validar acessibilidade, uso em tamanhos comuns e `prefers-reduced-motion`.
- [ ] Antes de produção: OAuth real, Graph consentido, controles de acesso, LGPD, persistência, CSRF, rate limiting, auditoria e backup/restore.

## Anexos efetivamente analisados

- `index.html`: estrutura SaaS com login, menu lateral, painel, cartões, painéis e formulários; branding existente é “Duarte Repay”, não foi reaproveitado literalmente.
- `Relatorio_Consolidado_20260629_1226.xlsx`: abas de resumo e dois detalhamentos de cliente; estados vazios/zeros indicados acima.
- `Mostrando-Sistema.md`: referência narrativa de três páginas com identificação, resumo, participantes, cronologia, tópicos, decisões, ações, riscos e próximos passos; nome da marca não se aplica.
- Matriz `.xlsm`: abas FLUXO OPERACIONAL, TRABALHOS DIÁRIOS, REUNIÕES E VISITAS, RESUMO OPERADORAS, DETALHES OPERADORAS, FICHA CADASTRAL, ESPECIALIDADE e CONTATOS. FLUXO OPERACIONAL tem cabeçalho COD, CLIENTE, ID, OPERADORA, CATEGORIA, ID PRIOR., PRIORIDADE, NÚMERO ANS, VIDAS NA ANS, PORTE, HONORÁRIOS, DIFERENCIAL e campos de fluxo/status; TRABALHOS DIÁRIOS tem DATA, OPERADORA/CLIENTE, DESCRIÇÃO DA ATIVIDADE e STATUS; REUNIÕES E VISITAS tem COD, CLIENTE, ID OPER., OPERADORA, DATA, HORÁRIO, TIPO, MOTIVO, STATUS, DETALHE, PARTICIPANTES, OBSERVAÇÃO; RESUMO/DETALHES contêm datas contratuais, prestador e dados de serviços; CONTATOS contém códigos, cliente, operadora e colunas de datas/contatos. Encontradas fórmulas `_xlfn.XLOOKUP` com referências externas. Proteção de workbook reportada como ausente pelo parser e nenhum mecanismo de senha utilizado; não alteramos o arquivo original.
#   d u a r t e r e l a t s  
 