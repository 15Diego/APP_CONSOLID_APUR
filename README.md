# Consolidador de Relatórios Base

[![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io)
[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?logo=python&logoColor=white)](https://www.python.org)

Aplicação web moderna para consolidação de múltiplas planilhas Excel em um único arquivo formatado.

## ✨ Funcionalidades

- 📤 **Upload múltiplo** de arquivos Excel (.xlsx)
- 🔍 **Detecção automática** de cabeçalhos
- 📊 **Consolidação inteligente** com alinhamento de colunas
- 🔗 **Rastreabilidade** completa (origem e linha de cada registro)
- 💅 **Formatação profissional** automática (tabelas, filtros, larguras)
- 📥 **Download** em Excel ou CSV
- 🌐 **Interface web** responsiva e moderna

## 🚀 Início Rápido

### Instalação

```bash
# Clone ou navegue até o diretório
cd 03_Consolidador_Apuração

# Instale as dependências
pip install -r requirements.txt
```

### Executar Localmente

```bash
streamlit run app.py
```

A aplicação abrirá automaticamente no navegador em `http://localhost:8501`

## 🎯 Como Usar

1. **Configure as opções** na barra lateral:
   - Nome da aba a consolidar
   - Detecção automática de cabeçalho (ou manual)
   - Opções avançadas (texto, auditoria, formatação)

2. **Selecione os arquivos** Excel para consolidar

3private. **Clique em "Consolidar Arquivos"**

4. **Visualize os resultados** nas abas:
   - Dados consolidados
   - Resumo do processamento

5. **Baixe o resultado** em Excel ou CSV

## 📁 Estrutura do Projeto

```
03_Consolidador_Apuração/
├── app.py                           # Aplicação Streamlit
├── consolidate_relatorio_base.py    # Lógica de consolidação (core)
├── requirements.txt                 # Dependências Python
└── README.md                        # Este arquivo
```

## 🛠️ Tecnologias

- **[Streamlit](https://streamlit.io)** - Framework web moderno para Python
- **[pandas](https://pandas.pydata.org/)** - Manipulação de dados
- **[openpyxl](https://openpyxl.readthedocs.io/)** - Leitura/escrita de Excel

## 📝 Versões

### v2.0 - Streamlit (Atual)
- Interface web moderna
- Upload direto de arquivos
- Visualização interativa de dados
- Download instantâneo

### v1.0 - Tkinter (Desktop)
- GUI desktop com Tkinter
- Disponível em `consolidate_relatorio_base.py`
- Execute com: `python consolidate_relatorio_base.py`

## 🎨 Capturas de Tela

> **Nota**: Execute a aplicação para ver a interface moderna e responsiva!

## 🔧 Opções Avançadas

### Ler como Texto
Preserva zeros à esquerda e evita conversões automáticas de tipos.

### Colunas de Auditoria
Adiciona informações de rastreabilidade:
- `ARQUIVO_ORIGEM` - Nome do arquivo de origem
- `ABA_ORIGEM` - Nome da aba lida
- `HEADER_LINHA` - Linha onde estava o cabeçalho
- `LINHA_ORIGEM_EXCEL` - Número da linha original no Excel

### Formatação Profissional
Aplica automaticamente:
- Tabelas formatadas do Excel
- Auto-filtro em todas as colunas
- Congelamento da linha de cabeçalho
- Ajuste automático de larguras de coluna

## 🚢 Deploy

### Streamlit Cloud (Gratuito)

1. Faça push do código para GitHub
2. Acesse [share.streamlit.io](https://share.streamlit.io)
3. Conecte seu repositório
4. Deploy automático!

### Outras Opções
- Heroku
- AWS (EC2, ECS)
- Google Cloud Run
- Azure App Service

## 📞 Suporte

Para problemas ou sugestões, consulte os logs da aplicação ou revise o código em `consolidate_relatorio_base.py`.

## 📄 Licença

Código interno - Uso restrito ao projeto.

---

**Desenvolvido com ❤️ usando Streamlit**
