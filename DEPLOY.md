# Implantação — Consolidador de Apuração

## Publicação no Streamlit Cloud

1. Confirme que o repositório `15Diego/APP_CONSOLID_APUR` está na branch `main`.
2. No Streamlit Cloud, selecione `app.py` como arquivo principal.
3. Após o push, use **Reboot app** para carregar a nova versão.

Esta versão usa acesso direto, sem login, cadastro, secrets ou conexão obrigatória com Supabase. Os arquivos processados, perfis e histórico ficam somente na sessão atual do navegador. Baixe os relatórios antes de fechar ou recarregar a página.

## Verificação pós-publicação

| Verificação | Resultado esperado |
|---|---|
| Tela inicial | Painel “Central de operações”, sem formulário de login |
| Consolidação | Pré-validação de aba, cabeçalho, linhas e erros por arquivo |
| Exportações | Download de Excel e CSV brasileiro |
| Histórico | Execuções disponíveis enquanto a sessão estiver aberta |
