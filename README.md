# Consolidador de Apuração

Aplicação Streamlit para pré-validar, consolidar, analisar e auditar planilhas fiscais. A versão atual possui **acesso direto**, sem login ou cadastro obrigatório.

## Funcionalidades

| Área | Recursos |
|---|---|
| Upload | `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls` e `.xlsb`, com validação de tamanho e prévia de estrutura |
| Leitura | Detecção automática de cabeçalho, busca tolerante de aba e alinhamento de colunas |
| Filtros | `CFOP`, `TES`, `Tipo de Movimento`, `Descrição do Produto`, período, filial e `UF` |
| Qualidade | Duplicidades, campos vazios, distribuição de `CFOP`/`TES`, totais e conferência por arquivo |
| Exportação | Excel profissional, CSV brasileiro com separador `;` e decimal `,`, e relatório de processamento |
| Sessão | Perfis como `AGLO – Relatório Base`, histórico e downloads enquanto a sessão estiver aberta |
| Acesso | Direto, sem login, cadastro ou senha |

## Arquitetura

O Streamlit executa o processamento diretamente na sessão do navegador. Por privacidade, arquivos e resultados não são persistidos após encerrar ou recarregar a sessão; faça download dos relatórios antes de sair.

## Configuração local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Não é necessário configurar secrets para usar esta versão.

## Deploy no Streamlit Cloud

1. Faça push deste repositório para o GitHub.
2. Confirme que o arquivo principal é `app.py` e reinicie o app após o push.

## Testes

```bash
python3 -m unittest discover tests
python3 -m py_compile app.py supabase_repository.py consolidate_relatorio_base.py
```

## Estrutura

```text
app.py                         # Interface Streamlit atual
app_legacy.py                  # Backup da interface anterior
consolidate_relatorio_base.py  # Núcleo de leitura, filtros e exportação Excel
supabase_repository.py         # Autenticação, armazenamento, histórico e auditoria
supabase_schema.sql            # Estrutura do banco Supabase
.streamlit/secrets.example.toml # Modelo de configuração confidencial
```
