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
    MAX_FILE_SIZE_MB,
    SUPPORTED_EXCEL_EXTENSIONS,
    SUPPORTED_EXCEL_EXTENSIONS_LABEL,
    AuditColumn,
    ler_planilha_robusta,
    consolidar_planilhas,
    salvar_excel,
    criar_excel_em_memoria,
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
    defaults = {
        "processed_files": [],
        "consolidated_df": None,
        "summary_df": None,
        "format_output": True,
        "sheet_name_used": DEFAULT_SHEET_NAME,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def clear_results() -> None:
    """Limpa resultados anteriores do session state."""
    st.session_state.consolidated_df = None
    st.session_state.summary_df = None
    st.session_state.processed_files = []


# ========================================
# Funções de Processamento
# ========================================

def validar_arquivos(uploaded_files: List) -> List[str]:
    """Valida extensão e tamanho dos arquivos enviados."""
    avisos = []
    for arquivo in uploaded_files:
        extensao = Path(arquivo.name).suffix.lower()
        if extensao not in SUPPORTED_EXCEL_EXTENSIONS:
            avisos.append(
                f"❌ **{arquivo.name}** possui formato não suportado. "
                f"Formatos aceitos: {SUPPORTED_EXCEL_EXTENSIONS_LABEL}."
            )
            continue

        size_mb = len(arquivo.getvalue()) / (1024 * 1024)
        if size_mb > MAX_FILE_SIZE_MB:
            avisos.append(
                f"⚠️ **{arquivo.name}** ({size_mb:.1f} MB) excede o limite recomendado de "
                f"{MAX_FILE_SIZE_MB} MB; o processamento pode demorar mais."
            )
    return avisos


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
    Processa arquivos usando BytesIO — sem escrita em disco.
    """
    results: List[ReadResult] = []
    dfs_ok: List[pd.DataFrame] = []

    progress_bar = st.progress(0)
    status_text = st.empty()
    header_0based = None if auto_detect_header else max(0, header_row - 1)

    for idx, uploaded_file in enumerate(uploaded_files, start=1):
        progress_bar.progress(idx / len(uploaded_files))
        status_text.text(f"Processando {idx}/{len(uploaded_files)}: {uploaded_file.name}")

        # Passa BytesIO diretamente — sem escrita em disco
        buffer = io.BytesIO(uploaded_file.getvalue())
        r = ler_planilha_robusta(
            file_path=buffer,
            preferred_sheet=preferred_sheet,
            auto_detect_header=auto_detect_header,
            header_row_0based=header_0based,
            read_as_text=read_as_text,
            adicionar_auditoria=add_audit,
            filtros=filtros,
            nome_arquivo=uploaded_file.name,
        )

        results.append(r)
        if r.status == "OK" and r.df is not None:
            dfs_ok.append(r.df)
            logger.info(f"✅ {r.arquivo}: {r.linhas} linhas")
        else:
            logger.error(f"❌ {r.arquivo}: {r.erro}")

    progress_bar.empty()
    status_text.empty()

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
    Cria arquivo Excel em memória para download.
    
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
    
    # Usa openpyxl (sem constant_memory) para garantir escrita correta de todos os dados
    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        df_dados.to_excel(writer, sheet_name=sheet_name, index=False)
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
        
        # Aplica formatação básica: congela linha do cabeçalho e ajusta larguras
        for ws_name in [sheet_name, "Resumo"]:
            ws = writer.sheets[ws_name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                max_len = max(
                    (len(str(cell.value)) if cell.value is not None else 0)
                    for cell in col
                )
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
    
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
        with st.expander("📝 Sobre"):
            st.markdown("""
            **Versão:** 2.3

            **Funcionalidades:**
            - ✅ Detecção automática de cabeçalho
            - ✅ Consolidação inteligente de múltiplas planilhas
            - ✅ Busca de aba tolerante (espaço, underscore, maiúsculas)
            - ✅ Rastreabilidade completa (arquivo, aba e linha de origem)
            - ✅ Formatação profissional no Excel (tabelas, filtros, larguras)
            - ✅ Upload de arquivos Excel: XLSX, XLSM, XLTX, XLTM, XLS e XLSB
            - ✅ Download em Excel e CSV (padrão brasileiro)
            - ✅ Filtros de extração: CFOP, TES, Tipo de Movimento e Descrição
            - ✅ Otimização de memória para arquivos grandes
            - ✅ Tratamento robusto de erros com mensagens detalhadas
            - ✅ Interface web moderna e responsiva
            """)
    
    return sheet_name, auto_detect, header_row, read_as_text, add_audit, format_output, filtros


def render_upload_section() -> tuple[List, bool, bool]:
    """
    Renderiza a seção de upload de arquivos.

    Returns:
        Tupla com (uploaded_files, process_button, clear_button)
    """
    st.subheader("📂 Upload de Arquivos")
    uploaded_files = st.file_uploader(
        "Selecione os arquivos Excel para consolidar",
        type=[ext.lstrip(".") for ext in SUPPORTED_EXCEL_EXTENSIONS],
        accept_multiple_files=True,
        help=(
            "Formatos aceitos: .xlsx, .xlsm, .xltx, .xltm, .xls e .xlsb. "
            "Arquivos com macros são lidos como dados; as macros não são executadas."
        )
    )

    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} arquivo(s) selecionado(s)")

        # Preview dos arquivos
        with st.expander("👁️ Visualizar arquivos selecionados"):
            for f in uploaded_files:
                file_size_mb = len(f.getvalue()) / (1024 * 1024)
                st.text(f"• {f.name} ({file_size_mb:.2f} MB)")

        # Validação de tamanho dos arquivos
        avisos = validar_arquivos(uploaded_files)
        if avisos:
            for aviso in avisos:
                st.warning(aviso)

    # Botões de ação abaixo da seção de upload
    col1, col2 = st.columns(2)

    with col1:
        process_button = st.button(
            "🔄 Consolidar",
            type="primary",
            disabled=not uploaded_files,
            use_container_width=True
        )

    with col2:
        clear_button = st.button(
            "🗑️ Limpar",
            use_container_width=True
        )

    return uploaded_files, process_button, clear_button


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
        # Preparar column_config para destacar colunas de auditoria
        df = st.session_state.consolidated_df
        column_config = {}

        # Destacar colunas de auditoria com background
        for col in df.columns:
            if col in ["arquivo_origem", "linha_origem", "aba_origem"]:
                column_config[col] = st.column_config.Column(
                    label=col,
                    help="Coluna de auditoria",
                    width="medium"
                )

        st.dataframe(
            df,
            height=400,
            use_container_width=True,
            column_config=column_config if column_config else None
        )

        # Estatísticas rápidas
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Linhas", f"{len(df):,}")
        with col2:
            st.metric("Total de Colunas", len(df.columns))
        with col3:
            memory_mb = df.memory_usage(deep=True).sum() / (1024 * 1024)
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
    uploaded_files, process_button, clear_button = render_upload_section()

    # Limpar resultados quando botão é clicado
    if clear_button:
        clear_results()
        st.rerun()

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

                # Exibe erros detalhados por arquivo imediatamente
                if fail_count > 0:
                    with st.expander(f"⚠️ {fail_count} arquivo(s) com erro — clique para ver detalhes", expanded=True):
                        for r in results:
                            if r.status != "OK":
                                st.error(
                                    f"📄 **{r.arquivo}**\n\n"
                                    f"{r.erro or 'Erro desconhecido.'}"
                                )

            except Exception as e:
                import traceback
                st.error(
                    f"❌ **Erro inesperado durante o processamento**\n\n"
                    f"**Tipo:** `{type(e).__name__}`\n\n"
                    f"**Mensagem:** {str(e)}"
                )
                with st.expander("🔍 Ver detalhes técnicos do erro"):
                    st.code(traceback.format_exc(), language="python")
                logger.error(f"Erro: {e}", exc_info=True)
    
    # Renderiza seção de resultados
    render_results_section(sheet_name)


if __name__ == "__main__":
    main()
