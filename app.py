"""
Aplicação Streamlit - Consolidador de Relatórios Base

Interface web moderna para consolidação de múltiplas planilhas Excel,
substituindo a interface Tkinter por uma solução web responsiva.
"""

import streamlit as st
import pandas as pd
import logging
from pathlib import Path
from typing import List, Optional
import io
from datetime import datetime
import gc
import shutil

# Import das funções do módulo principal
from consolidate_relatorio_base import (
    DEFAULT_SHEET_NAME,
    ler_planilha_robusta,
    consolidar_planilhas,
    salvar_excel,
    ReadResult,
    FilterConfig,
)


# ========================================
# Configuração da Página
# ========================================

st.set_page_config(
    page_title="Consolidador de Relatórios",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ========================================
# Configuração de Logging
# ========================================

@st.cache_resource
def setup_logger():
    """Configura logger para a aplicação."""
    logger = logging.getLogger("consolidacao_streamlit")
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        logger.addHandler(handler)
    
    return logger


logger = setup_logger()


# ========================================
# Session State
# ========================================

def init_session_state():
    """Inicializa variáveis de estado da sessão."""
    if "processed_files" not in st.session_state:
        st.session_state.processed_files = []
    if "consolidated_df" not in st.session_state:
        st.session_state.consolidated_df = None
    if "summary_df" not in st.session_state:
        st.session_state.summary_df = None


# ========================================
# Funções de Processamento
# ========================================

def process_uploaded_files(
    uploaded_files: List,
    preferred_sheet: str,
    auto_detect_header: bool,
    header_row: Optional[int],
    read_as_text: bool,
    add_audit: bool,
    filtros: Optional[FilterConfig] = None,
) -> tuple[pd.DataFrame, pd.DataFrame, List[ReadResult]]:
    """
    Processa arquivos enviados pelo usuário com tratamento robusto de erros.
    
    Args:
        uploaded_files: Lista de arquivos enviados via file_uploader.
        preferred_sheet: Nome da aba a ler.
        auto_detect_header: Se True, detecta cabeçalho automaticamente.
        header_row: Linha de cabeçalho manual (1-based).
        read_as_text: Se True, lê como texto.
        add_audit: Se True, adiciona colunas de auditoria.
        filtros: Configuração de filtros opcionais.
    
    Returns:
        Tupla (df_consolidado, df_resumo, resultados).
    """
    results: List[ReadResult] = []
    dfs_ok: List[pd.DataFrame] = []
    
    # Cria diretório temporário para processamento
    temp_dir = Path("temp_uploads")
    temp_dir.mkdir(exist_ok=True)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        for idx, uploaded_file in enumerate(uploaded_files, start=1):
            # Salva arquivo temporariamente
            temp_path = temp_dir / uploaded_file.name
            try:
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Atualiza progresso
                progress = idx / len(uploaded_files)
                progress_bar.progress(progress)
                status_text.text(f"Processando {idx}/{len(uploaded_files)}: {uploaded_file.name}")
                
                # Processa arquivo
                header_0based = None if auto_detect_header else max(0, header_row - 1)
                
                r = ler_planilha_robusta(
                    file_path=str(temp_path),
                    preferred_sheet=preferred_sheet,
                    auto_detect_header=auto_detect_header,
                    header_row_0based=header_0based,
                    read_as_text=read_as_text,
                    adicionar_auditoria=add_audit,
                    filtros=filtros,
                )
                
                results.append(r)
                
                if r.status == "OK" and r.df is not None:
                    dfs_ok.append(r.df)
                    logger.info(f"✅ {r.arquivo}: {r.linhas} linhas")
                else:
                    logger.error(f"❌ {r.arquivo}: {r.erro}")
            
            finally:
                # Remove arquivo temporário após processamento
                if temp_path.exists():
                    temp_path.unlink()
    
    finally:
        # Garante limpeza de diretório temporário mesmo em caso de erro
        try:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
        except Exception as cleanup_error:
            logger.warning(f"Não foi possível limpar diretório temporário: {cleanup_error}")
    
    # Consolida DataFrames
    df_consolidado = consolidar_planilhas(dfs_ok)

    # Otimização de Memória
    del dfs_ok
    for r in results:
        r.df = None
    gc.collect()
    
    # Gera resumo
    df_resumo = pd.DataFrame([
        {
            "arquivo": r.arquivo,
            "status": r.status,
            "aba": r.aba,
            "header_linha": (r.header_row_0based + 1) if r.header_row_0based is not None else None,
            "linhas": r.linhas,
            "colunas": r.colunas,
            "erro": r.erro,
        }
        for r in results
    ])
    
    progress_bar.empty()
    status_text.empty()
    
    return df_consolidado, df_resumo, results


def create_excel_download(df_dados: pd.DataFrame, df_resumo: pd.DataFrame, sheet_name: str) -> bytes:
    """
    Cria arquivo Excel em memória para download com otimização de memória.
    
    Args:
        df_dados: DataFrame com dados consolidados.
        df_resumo: DataFrame com resumo da consolidação.
        sheet_name: Nome da aba de dados.
    
    Returns:
        Bytes do arquivo Excel.
    """
    output = io.BytesIO()
    
    # Força garbage collection antes do processo pesado de escrita
    gc.collect()
    
    # Usa engine xlsxwriter com modo constant_memory para streaming
    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs={'options': {'constant_memory': True}}
    ) as writer:
        df_dados.to_excel(writer, sheet_name=sheet_name, index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
    
    return output.getvalue()


def create_csv_download_br(df: pd.DataFrame) -> str:
    """
    Cria CSV no padrão brasileiro (separador ; e decimal ,).
    
    Args:
        df: DataFrame a exportar.
    
    Returns:
        String CSV formatada no padrão brasileiro.
    """
    return df.to_csv(
        index=False,
        sep=';',
        decimal=',',
        encoding='utf-8-sig'
    )


def render_sidebar_config() -> tuple[str, bool, int, bool, bool, bool, FilterConfig]:
    """
    Renderiza a barra lateral com configurações e retorna os valores.
    
    Returns:
        Tupla com (sheet_name, auto_detect, header_row, read_as_text, add_audit, format_output, filtros)
    """
    with st.sidebar:
        st.header("⚙️ Configurações")
        
        # Nome da aba
        sheet_name = st.text_input(
            "Nome da aba",
            value=DEFAULT_SHEET_NAME,
            help="Nome da aba a ler em todos os arquivos"
        )
        
        # Detecção de cabeçalho
        auto_detect = st.checkbox(
            "Detectar cabeçalho automaticamente",
            value=True,
            help="Se ativado, detecta automaticamente a linha de cabeçalho"
        )
        
        header_row = 2
        if not auto_detect:
            header_row = st.number_input(
                "Linha do cabeçalho",
                min_value=1,
                max_value=100,
                value=2,
                help="Número da linha que contém o cabeçalho (1-based)"
            )
        
        st.divider()
        
        # Opções avançadas
        with st.expander("🔧 Opções Avançadas"):
            read_as_text = st.checkbox(
                "Ler tudo como texto",
                value=True,
                help="Preserva zeros à esquerda e evita conversões automáticas"
            )
            
            add_audit = st.checkbox(
                "Adicionar colunas de auditoria",
                value=True,
                help="Adiciona colunas com arquivo de origem e número da linha"
            )
            
            format_output = st.checkbox(
                "Formatar saída",
                value=True,
                help="Aplica formatação profissional (tabelas, filtros, larguras)"
            )
            
        with st.expander("🔍 Filtros de Extração"):
            st.caption("Deixe em branco para não filtrar.")
            
            f_cfop = st.text_input(
                "CFOP",
                placeholder="Ex: 5102, 6102",
                help="Códigos separados por vírgula"
            )
            
            f_tes = st.text_input(
                "TES",
                placeholder="Ex: 501, 502",
                help="Códigos separados por vírgula"
            )
            
            f_tm = st.text_input(
                "Tipo de Movimento",
                placeholder="Ex: VENDA, DEVOLUCAO",
                help="Tipos separados por vírgula"
            )
            
            f_desc = st.text_input(
                "Descrição do Produto",
                placeholder="Ex: PARAFUSO",
                help="Filtrar produtos que contenham este termo"
            )
            
            # Cria objeto de configuração de filtros
            filtros = FilterConfig(
                cfops=[x.strip() for x in f_cfop.split(",") if x.strip()],
                tes=[x.strip() for x in f_tes.split(",") if x.strip()],
                tipo_movimento=[x.strip() for x in f_tm.split(",") if x.strip()],
                descricao_contem=f_desc.strip()
            )
        
        st.divider()
        st.markdown("### 📝 Sobre")
        st.markdown("""
        Versão: **2.1** (Streamlit - Melhorado)
        
        Funcionalidades:
        - ✅ Detecção automática de cabeçalho
        - ✅ Consolidação inteligente
        - ✅ Rastreabilidade completa
        - ✅ Formatação profissional
        - ✅ Interface web moderna
        - ✅ Tratamento robusto de erros
        """)
    
    return sheet_name, auto_detect, header_row, read_as_text, add_audit, format_output, filtros


def render_upload_section() -> List:
    """
    Renderiza a seção de upload de arquivos.
    
    Returns:
        Lista de arquivos enviados.
    """
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.subheader("📂 Upload de Arquivos")
        uploaded_files = st.file_uploader(
            "Selecione os arquivos Excel (.xlsx) para consolidar",
            type=["xlsx"],
            accept_multiple_files=True,
            help="Carregue múltiplos arquivos Excel para consolidação"
        )
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} arquivo(s) selecionado(s)")
            
            # Preview dos arquivos
            with st.expander("👁️ Visualizar arquivos selecionados"):
                for f in uploaded_files:
                    file_size_mb = len(f.getvalue()) / (1024 * 1024)
                    st.text(f"• {f.name} ({file_size_mb:.2f} MB)")
    
    with col2:
        st.subheader("🚀 Ação")
        process_button = st.button(
            "🔄 Consolidar",
            type="primary",
            disabled=not uploaded_files,
            use_container_width=True
        )
    
    return uploaded_files, process_button


def render_results_section(sheet_name: str):
    """
    Renderiza a seção de resultados com dados consolidados e download.
    
    Args:
        sheet_name: Nome da aba de dados.
    """
    if st.session_state.consolidated_df is None:
        return
    
    st.divider()
    st.subheader("📊 Resultados")
    
    tab1, tab2 = st.tabs(["📈 Dados Consolidados", "📋 Resumo"])
    
    with tab1:
        st.dataframe(
            st.session_state.consolidated_df,
            height=400,
            use_container_width=True
        )
        
        # Estatísticas rápidas
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Linhas", f"{len(st.session_state.consolidated_df):,}")
        with col2:
            st.metric("Total de Colunas", len(st.session_state.consolidated_df.columns))
        with col3:
            memory_mb = st.session_state.consolidated_df.memory_usage(deep=True).sum() / (1024 * 1024)
            st.metric("Memória", f"{memory_mb:.2f} MB")
    
    with tab2:
        st.dataframe(
            st.session_state.summary_df,
            height=400,
            use_container_width=True
        )
    
    # Download
    st.divider()
    st.subheader("💾 Download")
    
    col1, col2 = st.columns(2)
    
    with col1:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"consolidado_{timestamp}.xlsx"
        
        excel_bytes = create_excel_download(
            st.session_state.consolidated_df,
            st.session_state.summary_df,
            sheet_name
        )
        
        st.download_button(
            label="📥 Download Excel",
            data=excel_bytes,
            file_name=default_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    
    with col2:
        csv_data = create_csv_download_br(st.session_state.consolidated_df)
        st.download_button(
            label="📥 Download CSV (BR)",
            data=csv_data,
            file_name=f"consolidado_{timestamp}.csv",
            mime="text/csv",
            help="CSV com separador ; e decimal , (padrão brasileiro)",
            use_container_width=True
        )


# ========================================
# Interface Principal
# ========================================

def main():
    """Função principal da aplicação Streamlit."""
    
    init_session_state()
    
    # Header
    st.title("📊 Consolidador de Relatórios Base")
    st.markdown("**Consolide múltiplas planilhas Excel em um único arquivo formatado**")
    st.divider()
    
    # Sidebar - Configurações
    sheet_name, auto_detect, header_row, read_as_text, add_audit, format_output, filtros = render_sidebar_config()
    
    # Main content
    uploaded_files, process_button = render_upload_section()
    
    st.divider()
    
    # Processamento
    if process_button and uploaded_files:
        with st.spinner("Processando arquivos..."):
            try:
                df_consolidated, df_summary, results = process_uploaded_files(
                    uploaded_files=uploaded_files,
                    preferred_sheet=sheet_name,
                    auto_detect_header=auto_detect,
                    header_row=header_row,
                    read_as_text=read_as_text,
                    add_audit=add_audit,
                    filtros=filtros,
                )
                
                # Salva em session state
                st.session_state.consolidated_df = df_consolidated
                st.session_state.summary_df = df_summary
                st.session_state.processed_files = results
                
                # Estatísticas
                ok_count = sum(r.status == "OK" for r in results)
                fail_count = sum(r.status != "OK" for r in results)
                
                if df_consolidated.empty:
                    st.error("❌ Nenhum arquivo foi consolidado com sucesso. Verifique o resumo abaixo para detalhes dos erros.")
                else:
                    st.success(f"""
                    ✅ **Consolidação concluída!**
                    
                    - Arquivos processados com sucesso: {ok_count}/{len(results)}
                    - Arquivos com falha: {fail_count}
                    - Linhas consolidadas: {len(df_consolidated):,}
                    - Colunas: {len(df_consolidated.columns)}
                    """)
                
            except Exception as e:
                st.error(f"❌ Erro durante processamento: {str(e)}")
                logger.error(f"Erro: {e}", exc_info=True)
    
    # Renderiza seção de resultados
    render_results_section(sheet_name)


if __name__ == "__main__":
    main()
