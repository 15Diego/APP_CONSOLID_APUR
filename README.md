# Consolidador de Apuração

Aplicação Streamlit para pré-validar, consolidar, analisar e auditar planilhas fiscais. A versão atual usa **Supabase** para autenticação, perfis reutilizáveis, histórico de processamentos, arquivos privados e trilha de auditoria.

## Funcionalidades

| Área | Recursos |
|---|---|
| Upload | `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls` e `.xlsb`, com validação de tamanho e prévia de estrutura |
| Leitura | Detecção automática de cabeçalho, busca tolerante de aba e alinhamento de colunas |
| Filtros | `CFOP`, `TES`, `Tipo de Movimento`, `Descrição do Produto`, período, filial e `UF` |
| Qualidade | Duplicidades, campos vazios, distribuição de `CFOP`/`TES`, totais e conferência por arquivo |
| Exportação | Excel profissional, CSV brasileiro com separador `;` e decimal `,`, e relatório de processamento |
| Persistência | Perfis como `AGLO – Relatório Base`, histórico, auditoria e downloads privados |
| Acesso | Exatamente dois perfis: `admin` e `usuário` |

## Arquitetura

O Streamlit executa o processamento das planilhas e usa o Supabase para autenticação, banco de dados e armazenamento privado. O primeiro usuário registrado recebe o perfil `admin`; os demais recebem `usuário` e podem ter seu perfil ajustado pela tela **Administração**.

## Configuração local

```bash
pip install -r requirements.txt
cp .streamlit/secrets.example.toml .streamlit/secrets.toml
streamlit run app.py
```

No arquivo `.streamlit/secrets.toml`, informe a `SUPABASE_SERVICE_ROLE_KEY` do projeto. Esse arquivo é confidencial e não deve ser enviado ao GitHub.

## Deploy no Streamlit Cloud

1. Faça push deste repositório para o GitHub.
2. No Streamlit Cloud, abra **App settings → Secrets**.
3. Copie o conteúdo de `.streamlit/secrets.example.toml` e substitua o valor da `SUPABASE_SERVICE_ROLE_KEY` pela chave correspondente do Supabase.
4. Confirme que o arquivo principal é `app.py` e reinicie o app.

> A chave de serviço deve ficar somente nos secrets do Streamlit Cloud. Ela nunca deve ser colocada no código, em commits ou em variáveis expostas ao navegador.

## Banco de dados Supabase

O arquivo `supabase_schema.sql` contém a estrutura de usuários, perfis, processamentos, arquivos, auditoria e regras de segurança. A migração já foi aplicada no projeto Supabase associado.

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
