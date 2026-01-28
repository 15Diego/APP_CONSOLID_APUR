# Guia de Deploy - Consolidador de Relatórios

Este guia detalha o processo completo para colocar a aplicação em produção.

## 🚀 Opção 1: Streamlit Cloud (Recomendado)

**Vantagens**:
- ✅ **Gratuito** para projetos públicos
- ✅ Deploy em **minutos**
- ✅ **HTTPS** automático
- ✅ Atualizações automáticas via Git
- ✅ Sem necessidade de servidor próprio

### Pré-requisitos
- Conta no GitHub (gratuita)
- Conta no Streamlit Cloud (gratuita)

### Passo 1: Criar Repositório no GitHub

1. Acesse https://github.com/new
2. Configure:
   - **Nome**: `consolidador-relatorios`
   - **Descrição**: `Aplicação web para consolidação de planilhas Excel`
   - **Visibilidade**: Privado (se dados sensíveis) ou Público
3. Clique em **"Create repository"**

### Passo 2: Fazer Push do Código

Abra o terminal no diretório do projeto e execute:

```bash
cd "c:\Users\diego.silva\.vscode\0-Projetos_Diego\03_Consolidador_Apuraçao"

# Inicializar Git (se ainda não foi)
git init

# Adicionar arquivos
git add app.py
git add consolidate_relatorio_base.py
git add requirements.txt
git add README.md
git add .streamlit/config.toml
git add .gitignore

# Commit
git commit -m "Initial commit: Consolidador de Relatórios v2.0"

# Conectar ao repositório remoto
git remote add origin https://github.com/SEU_USUARIO/consolidador-relatorios.git

# Push para GitHub
git branch -M main
git push -u origin main
```

### Passo 3: Deploy no Streamlit Cloud

1. Acesse https://share.streamlit.io
2. Clique em **"New app"**
3. Configure:
   - **Repository**: Selecione `consolidador-relatorios`
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. Clique em **"Deploy!"**

**Aguarde 2-3 minutos** para o deploy completar.

### Passo 4: Acessar Aplicação

Você receberá um URL público, algo como:
```
https://seu-usuario-consolidador-relatorios-xxxxx.streamlit.app
```

**Pronto!** Sua aplicação está em produção! 🎉

---

## 🐳 Opção 2: Docker

### Passo 1: Criar Dockerfile

Crie um arquivo `Dockerfile`:

```dockerfile
FROM python:3.10-slim

WORKDIR /app

# Instalar dependências do sistema (se necessário)
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código
COPY app.py .
COPY consolidate_relatorio_base.py .
COPY .streamlit .streamlit

# Expor porta
EXPOSE 8501

# Healthcheck
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

# Comando para iniciar
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

### Passo 2: Build e Run

```bash
# Build da imagem
docker build -t consolidador-relatorios .

# Executar container
docker run -p 8501:8501 consolidador-relatorios
```

Acesse em: http://localhost:8501

### Passo 3: Deploy (Docker Hub + Cloud Provider)

```bash
# Tag e push para Docker Hub
docker tag consolidador-relatorios seu-usuario/consolidador-relatorios:latest
docker push seu-usuario/consolidador-relatorios:latest

# Deploy em qualquer cloud que suporte Docker
# Ex: AWS ECS, Google Cloud Run, Azure Container Instances
```

---

## ☁️ Opção 3: Heroku

### Passo 1: Criar Procfile

```bash
web: sh setup.sh && streamlit run app.py
```

### Passo 2: Criar setup.sh

```bash
mkdir -p ~/.streamlit/

echo "\
[server]\n\
headless = true\n\
port = $PORT\n\
enableCORS = false\n\
\n\
" > ~/.streamlit/config.toml
```

### Passo 3: Deploy

```bash
# Instalar Heroku CLI: https://devcenter.heroku.com/articles/heroku-cli

# Login
heroku login

# Criar app
heroku create consolidador-relatorios

# Push
git push heroku main

# Abrir
heroku open
```

---

## 🔐 Segurança em Produção

### Limite de Upload

Edite `.streamlit/config.toml`:

```toml
[server]
maxUploadSize = 200  # MB
maxMessageSize = 200  # MB
```

### Autenticação (Opcional)

Para adicionar login, instale:

```bash
pip install streamlit-authenticator
```

E adicione no `app.py`:

```python
import streamlit_authenticator as stauth

# Configurar usuários
names = ['Admin User']
usernames = ['admin']
passwords = ['senha_hash_aqui']

authenticator = stauth.Authenticate(
    names, usernames, passwords,
    'cookie_name', 'signature_key', cookie_expiry_days=30
)

name, authentication_status, username = authenticator.login('Login', 'main')

if authentication_status:
    main()  # Sua aplicação
elif authentication_status == False:
    st.error('Usuário/senha incorretos')
else:
    st.warning('Por favor, faça login')
```

### Secrets Management

Para dados sensíveis, use **Streamlit Secrets**:

1. No Streamlit Cloud: Settings → Secrets
2. Adicione variáveis no formato TOML:

```toml
[database]
host = "seu-host"
password = "sua-senha"
```

3. Acesse no código:

```python
import streamlit as st
db_host = st.secrets["database"]["host"]
```

---

## 📊 Monitoramento

### Logs no Streamlit Cloud

1. Acesse sua app no Streamlit Cloud
2. Clique em **"Manage app"**
3. Veja logs em tempo real

### Analytics (Opcional)

Adicione Google Analytics:

```python
# No app.py
import streamlit.components.v1 as components

ga_code = """
<!-- Global site tag (gtag.js) - Google Analytics -->
<script async src="https://www.googletagmanager.com/gtag/js?id=GA_MEASUREMENT_ID"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'GA_MEASUREMENT_ID');
</script>
"""

components.html(ga_code, height=0)
```

---

## 🔄 Atualizações

### Com Streamlit Cloud
1. Faça alterações no código local
2. Commit e push para GitHub:
   ```bash
   git add .
   git commit -m "Descrição da mudança"
   git push
   ```
3. **Atualização automática** no Streamlit Cloud!

### Com Docker
1. Rebuild a imagem
2. Push para registry
3. Restart do container

---

## 🐛 Troubleshooting

### Problema: App não inicia

**Solução**: Verifique logs
- Streamlit Cloud: Menu → View logs
- Local: Terminal onde rodou `streamlit run`

### Problema: Erro de memória

**Solução**: Aumente recursos ou otimize processamento
```python
# Processar em chunks menores
for chunk in pd.read_excel(file, chunksize=1000):
    # processar chunk
```

### Problema: Upload muito lento

**Solução**: Verifique tamanho máximo e compressão
- Reduza `maxUploadSize`
- Peça usuários para compactar arquivos grandes

---

## ✅ Checklist de Deploy

Antes de colocar em produção, verifique:

- [ ] Testado localmente com dados reais
- [ ] Tratamento de erros implementado
- [ ] Mensagens de erro claras para usuário
- [ ] Limites de upload configurados
- [ ] README.md atualizado
- [ ] .gitignore configurado (não enviar secrets)
- [ ] Logs configurados
- [ ] Performance testada com arquivos grandes
- [ ] Responsividade testada (mobile/desktop)
- [ ] Backup dos dados configurado (se aplicável)

---

## 🎯 Recomendação Final

**Para este projeto, recomendo: Streamlit Cloud**

**Por quê?**
- ✅ Setup em 5 minutos
- ✅ Totalmente gratuito
- ✅ HTTPS automático
- ✅ Atualizações via Git push
- ✅ Perfeito para uso interno/corporativo
- ✅ Escalável conforme necessidade

**Próximos passos**:
1. Criar repositório GitHub (privado para dados corporativos)
2. Push do código
3. Deploy no Streamlit Cloud
4. Compartilhar URL com equipe

---

## 📞 Suporte

Documentação oficial:
- **Streamlit Cloud**: https://docs.streamlit.io/streamlit-community-cloud
- **Streamlit**: https://docs.streamlit.io
- **Deploy Guides**: https://docs.streamlit.io/streamlit-community-cloud/get-started

---

**Tempo estimado para deploy: 10 minutos** ⏱️
