# Implantação — Consolidador de Apuração v3

## Pré-requisitos

O repositório deve estar conectado ao Streamlit Cloud e o projeto Supabase `DLS_BASE` deve estar ativo. A estrutura de dados, as políticas de segurança e o bucket privado `fiscal-files` já foram criados no Supabase.

## Configuração de secrets

No Streamlit Cloud, abra **App settings → Secrets** e configure:

```toml
SUPABASE_URL = "https://wjtmjcntawexilozbfke.supabase.co"
SUPABASE_SERVICE_ROLE_KEY = "sua_service_role_key_do_supabase"
```

A `SUPABASE_SERVICE_ROLE_KEY` deve ser copiada de **Supabase → Project Settings → API**. Nunca a inclua em commit, issue, mensagem ou arquivo público.

## Publicação

1. Confirme que o repositório `15Diego/APP_CONSOLID_APUR` está na branch `main`.
2. No Streamlit Cloud, selecione `app.py` como arquivo principal.
3. Salve os secrets e selecione **Reboot app** após o push.
4. No primeiro acesso, crie a primeira conta. Ela será registrada como `admin`; novas contas recebem o perfil `usuário`.

## Verificação pós-publicação

| Verificação | Resultado esperado |
|---|---|
| Tela inicial | Login ou criação de conta do Supabase |
| Consolidação | Pré-validação de aba, cabeçalho, linhas e erros por arquivo |
| Exportações | Download de Excel, CSV brasileiro e relatório de processamento |
| Histórico | Execuções e links privados de download |
| Administração | Visível apenas ao perfil `admin` |

## Segurança operacional

Os arquivos são armazenados em bucket privado e acessados por URLs temporárias. Mantenha o projeto/repositório privado, use senhas fortes e revise os usuários administrativos periodicamente.
