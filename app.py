"""Consolidador de Apuração — experiência Streamlit com Supabase."""

from __future__ import annotations

import io
import json
from datetime import datetime
from typing import Any, Iterable

import pandas as pd
import streamlit as st

from consolidate_relatorio_base import (
    DEFAULT_SHEET_NAME,
    MAX_FILE_SIZE_MB,
    SUPPORTED_EXCEL_EXTENSIONS,
    FilterConfig,
    criar_excel_em_memoria,
    consolidar_planilhas,
    ler_planilha_robusta,
)
import supabase_repository as repository


st.set_page_config(page_title="Consolidador de Apuração", page_icon="CA", layout="wide", initial_sidebar_state="expanded")


def inject_style() -> None:
    st.markdown(
        """
        <style>
        :root { --ink:#10172b; --muted:#667085; --line:#e7eaf1; --panel:#ffffff; --wash:#f6f7fb; --violet:#5b5ce2; --navy:#081126; }
        .stApp { background: var(--wash); color:var(--ink); }
        [data-testid="stSidebar"] { background:#fff; border-right:1px solid var(--line); }
        [data-testid="stSidebar"] > div:first-child { padding-top:1.2rem; }
        h1,h2,h3 { letter-spacing:-0.035em; color:var(--ink); }
        .hero { background:linear-gradient(125deg,#070d20 0%,#111c3c 58%,#463ca5 160%); color:#fff; border-radius:24px; padding:2.6rem 2.8rem; margin:0.2rem 0 1.7rem; box-shadow:0 18px 50px rgba(15,23,54,.17); }
        .hero .eyebrow { color:#d8d9ff; font-size:.78rem; font-weight:700; letter-spacing:.13em; text-transform:uppercase; }
        .hero h1 { color:#fff; font-size:2.45rem; margin:.5rem 0 .65rem; max-width:680px; line-height:1.08; }
        .hero p { color:#cbd5ef; font-size:1.03rem; margin:0; max-width:650px; }
        .metric-card { background:#fff; border:1px solid var(--line); border-radius:17px; padding:1.15rem 1.25rem; min-height:128px; box-shadow:0 2px 4px rgba(16,24,40,.02); }
        .metric-card .label { color:#7a8495; font-weight:700; font-size:.72rem; letter-spacing:.11em; text-transform:uppercase; }
        .metric-card .value { font-size:2.05rem; font-weight:750; letter-spacing:-.06em; margin:.45rem 0 .1rem; color:#17203b; }
        .metric-card .caption { color:#7a8495; font-size:.84rem; }
        .section-card { background:#fff; border:1px solid var(--line); border-radius:17px; padding:1.35rem; margin-bottom:1rem; }
        .section-title { font-size:1.1rem; font-weight:750; color:#1b2541; margin-bottom:.25rem; }
        .section-subtitle { color:#6d778a; font-size:.9rem; margin-bottom:1.1rem; }
        .status-ok { color:#078a5b; font-weight:700; } .status-error { color:#c93753; font-weight:700; }
        .stButton > button, .stDownloadButton > button { border-radius:10px; font-weight:650; min-height:2.65rem; }
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] { background:#5859dc; border-color:#5859dc; }
        .stTabs [data-baseweb="tab-list"] { gap:18px; border-bottom:1px solid var(--line); }
        .stTabs [data-baseweb="tab"] { font-weight:650; padding:11px 2px; }
        .stDataFrame { border:1px solid var(--line); border-radius:12px; overflow:hidden; }
        .auth-shell { max-width:1000px; margin:4vh auto; }
        .auth-card { background:#fff; border:1px solid var(--line); border-radius:20px; padding:2rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_state() -> None:
    defaults = {
        "auth": None,
        "profile": None,
        "prevalidation": [],
        "result_df": None,
        "summary_df": None,
        "quality": None,
        "last_job": None,
        "active_nav": "Visão geral",
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def canonical(value: Any) -> str:
    return clean_text(value).lower().replace("_", " ").replace(".", "")


def find_column(df: pd.DataFrame, names: Iterable[str]) -> str | None:
    items = [(canonical(col), col) for col in df.columns]
    for wanted in names:
        query = canonical(wanted)
        for normalised, original in items:
            if normalised == query:
                return original
        for normalised, original in items:
            if query in normalised:
                return original
    return None


def parse_number(value: Any) -> float:
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    raw = clean_text(value).replace("R$", "").replace(" ", "")
    if not raw:
        return 0.0
    if "," in raw and "." in raw:
        raw = raw.replace(".", "").replace(",", ".") if raw.rfind(",") > raw.rfind(".") else raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(".", "").replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return 0.0


def quality_summary(df: pd.DataFrame) -> dict[str, Any]:
    if df.empty:
        return {"records": 0, "duplicates": 0, "empty": [], "cfop": pd.DataFrame(), "tes": pd.DataFrame(), "total": 0.0}
    candidates = {"CFOP": ["CFOP", "C.F.O.P.", "CFOP Fiscal"], "TES": ["TES", "T.E.S"], "Tipo de Movimento": ["Tp. Mov", "Tipo de Movimento"], "Descrição do Produto": ["Desc. Produto", "Descrição do Produto"]}
    empty = []
    for label, names in candidates.items():
        column = find_column(df, names)
        empty.append({"Campo": label, "Vazios": int(df[column].isna().sum() + (df[column].astype(str).str.strip() == "").sum()) if column else len(df)})
    key_cols = [column for column in [find_column(df, ["Chave Doc", "Chave"]), find_column(df, ["Documento", "Doc"]), find_column(df, ["Serie", "Série"]), find_column(df, ["Filial"])] if column]
    duplicates = int(df.duplicated(subset=key_cols, keep="first").sum()) if key_cols else int(df.duplicated().sum())
    total_col = find_column(df, ["Vlr Contabil", "Valor Contabil", "Valor Total"])
    total = float(df[total_col].map(parse_number).sum()) if total_col else 0.0
    def distribution(names: list[str]) -> pd.DataFrame:
        col = find_column(df, names)
        if not col:
            return pd.DataFrame(columns=["Valor", "Registros"])
        return df[col].astype(str).str.strip().value_counts().head(12).rename_axis("Valor").reset_index(name="Registros")
    return {"records": len(df), "duplicates": duplicates, "empty": pd.DataFrame(empty), "cfop": distribution(candidates["CFOP"]), "tes": distribution(candidates["TES"]), "total": total}


def parse_mapping(text: str) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for row in text.splitlines():
        if "=>" not in row:
            continue
        source, target = (part.strip() for part in row.split("=>", 1))
        if source and target:
            mapping[source] = target
    return mapping


def available_values(prevalidation: list[dict[str, Any]], field: str) -> list[str]:
    return sorted({item for report in prevalidation for item in report.get("values", {}).get(field, []) if item})


def apply_filters(df: pd.DataFrame, filters: dict[str, Any]) -> pd.DataFrame:
    work = df.copy()
    choices = {
        "cfop": ["CFOP", "C.F.O.P.", "CFOP Fiscal"],
        "tes": ["TES", "T.E.S"],
        "tipo_movimento": ["Tp. Mov", "Tipo de Movimento", "Movimento"],
        "filial": ["Filial", "Loja", "Estabelecimento"],
        "uf": ["UF", "Estado"],
    }
    for key, names in choices.items():
        selected = filters.get(key, [])
        column = find_column(work, names)
        if selected and column:
            accepted = {canonical(item) for item in selected}
            work = work[work[column].map(canonical).isin(accepted)]
    description = filters.get("descricao", "")
    description_col = find_column(work, ["Desc. Produto", "Descrição do Produto", "Descricao Produto", "Descrição"])
    if description and description_col:
        work = work[work[description_col].astype(str).str.contains(description, case=False, na=False)]
    date_col = find_column(work, ["Data", "Dt. Entrada", "Dt. Emissao", "Data Emissao"])
    start, end = filters.get("periodo_inicial"), filters.get("periodo_final")
    if date_col and (start or end):
        dates = pd.to_datetime(work[date_col], errors="coerce", dayfirst=True)
        if start:
            work = work[dates >= pd.to_datetime(start)]
            dates = dates.loc[work.index]
        if end:
            work = work[dates <= pd.to_datetime(end)]
    return work


def login_screen() -> None:
    st.markdown("<div class='auth-shell'><div class='hero'><div class='eyebrow'>Ambiente fiscal protegido</div><h1>Consolidação com controle, clareza e rastreabilidade.</h1><p>Acesse a operação fiscal para processar, auditar e recuperar seus relatórios de forma segura.</p></div></div>", unsafe_allow_html=True)
    if not repository.configured():
        st.error("O Supabase ainda não foi configurado nos secrets do Streamlit Cloud. Consulte a seção de implantação no README.")
        st.stop()
    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("<div class='auth-card'>", unsafe_allow_html=True)
        st.subheader("Entrar")
        with st.form("login"):
            email = st.text_input("E-mail", key="login_email")
            password = st.text_input("Senha", type="password", key="login_password")
            submit = st.form_submit_button("Acessar ambiente", type="primary", use_container_width=True)
        if submit:
            try:
                st.session_state.auth = repository.sign_in(email, password)
                st.session_state.profile = repository.get_user_profile(st.session_state.auth["id"])
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
        st.markdown("</div>", unsafe_allow_html=True)
    with right:
        st.markdown("<div class='auth-card'>", unsafe_allow_html=True)
        st.subheader("Criar acesso")
        with st.form("signup"):
            name = st.text_input("Nome", key="signup_name")
            email = st.text_input("E-mail", key="signup_email")
            password = st.text_input("Senha", type="password", key="signup_password", help="Use ao menos 8 caracteres.")
            submit = st.form_submit_button("Criar conta", use_container_width=True)
        if submit:
            try:
                repository.sign_up(email, password, name)
                st.success("Conta criada. Confirme o e-mail, se solicitado, e depois entre no ambiente.")
            except Exception as exc:
                st.error(str(exc))
        st.markdown("</div>", unsafe_allow_html=True)


def sidebar(profile: dict[str, Any]) -> str:
    role_label = "Administrador" if profile.get("role") == "admin" else "Usuário"
    with st.sidebar:
        st.markdown("### Consolidador")
        st.caption("Apuração fiscal · ambiente protegido")
        nav = ["Visão geral", "Consolidar", "Perfis", "Histórico", "Relatórios"]
        if profile.get("role") == "admin":
            nav.append("Administração")
        current = st.radio("Navegação", nav, label_visibility="collapsed", key="main_nav")
        st.divider()
        st.markdown(f"**{profile.get('display_name') or profile.get('email')}**")
        st.caption(role_label)
        if st.button("Sair", use_container_width=True):
            st.session_state.auth = None
            st.session_state.profile = None
            st.rerun()
    return current


def dashboard(user_id: str) -> None:
    jobs = repository.list_jobs(user_id)
    rows = sum(int(job.get("total_rows") or 0) for job in jobs)
    files = sum(int(job.get("total_files") or 0) for job in jobs)
    st.markdown("<div class='hero'><div class='eyebrow'>Central de operações</div><h1>Consolidação fiscal com segurança e rastreabilidade.</h1><p>Pré-valide arquivos, padronize colunas, aplique filtros e mantenha o histórico dos relatórios produzidos.</p></div>", unsafe_allow_html=True)
    a, b, c = st.columns(3)
    for column, label, value, caption in [(a, "Processamentos", len(jobs), "execuções registradas"), (b, "Arquivos analisados", files, "no histórico"), (c, "Linhas consolidadas", rows, "com auditoria")]:
        with column:
            st.markdown(f"<div class='metric-card'><div class='label'>{label}</div><div class='value'>{value:,}</div><div class='caption'>{caption}</div></div>", unsafe_allow_html=True)
    st.markdown("<div class='section-card'><div class='section-title'>Fluxo de trabalho</div><div class='section-subtitle'>1. Pré-validar arquivos · 2. Padronizar e filtrar · 3. Consolidar e exportar · 4. Consultar o histórico auditável.</div></div>", unsafe_allow_html=True)


def prevalidate(files: list[Any], sheet_name: str, auto_header: bool, header_line: int) -> list[dict[str, Any]]:
    reports = []
    for file in files:
        result = ler_planilha_robusta(io.BytesIO(file.getvalue()), sheet_name, auto_header, None if auto_header else header_line - 1, True, False, None, file.name)
        values: dict[str, list[str]] = {"cfop": [], "tes": [], "tipo_movimento": [], "filial": [], "uf": []}
        if result.status == "OK" and result.df is not None:
            for key, names in {"cfop": ["CFOP", "C.F.O.P.", "CFOP Fiscal"], "tes": ["TES"], "tipo_movimento": ["Tp. Mov", "Tipo de Movimento"], "filial": ["Filial"], "uf": ["UF", "Estado"]}.items():
                col = find_column(result.df, names)
                if col:
                    values[key] = sorted(result.df[col].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())[:200]
        reports.append({"arquivo": file.name, "status": result.status, "aba": result.aba, "cabecalho": (result.header_row_0based + 1 if result.header_row_0based is not None else None), "linhas": result.linhas, "colunas": result.colunas, "erro": result.erro, "values": values})
    return reports


def render_quality(quality: dict[str, Any], per_file: pd.DataFrame) -> None:
    st.markdown("<div class='section-card'><div class='section-title'>Qualidade dos dados</div><div class='section-subtitle'>Indicadores gerados sobre o resultado consolidado e os arquivos de origem.</div>", unsafe_allow_html=True)
    a, b, c = st.columns(3)
    a.metric("Registros", f"{quality['records']:,}")
    b.metric("Duplicidades", f"{quality['duplicates']:,}")
    c.metric("Total contábil", f"R$ {quality['total']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    x, y = st.columns(2)
    with x:
        st.caption("Distribuição de CFOP")
        st.dataframe(quality["cfop"], use_container_width=True, hide_index=True)
    with y:
        st.caption("Distribuição de TES")
        st.dataframe(quality["tes"], use_container_width=True, hide_index=True)
    st.caption("Campos para revisão")
    st.dataframe(quality["empty"], use_container_width=True, hide_index=True)
    st.caption("Conferência por arquivo")
    st.dataframe(per_file, use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)


def consolidar_screen(user_id: str) -> None:
    st.title("Consolidar")
    st.caption("Envie arquivos, confirme a estrutura identificada e produza relatórios fiscalmente rastreáveis.")
    with st.expander("Configuração da leitura", expanded=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        sheet_name = c1.text_input("Nome da aba", value=DEFAULT_SHEET_NAME)
        auto_header = c2.checkbox("Detectar cabeçalho", value=True)
        header_line = int(c3.number_input("Linha manual", min_value=1, value=2, disabled=auto_header))
        mapping_text = st.text_area("Mapeamento de colunas", placeholder="CFOP Fiscal => CFOP\nC.F.O.P. => CFOP\nTp. Mov => Tipo de Movimento", help="Uma equivalência por linha, no formato coluna de origem => nome padronizado.")
    files = st.file_uploader("Arquivos Excel", type=[item.lstrip(".") for item in SUPPORTED_EXCEL_EXTENSIONS], accept_multiple_files=True, help="Formatos aceitos: .xlsx, .xlsm, .xltx, .xltm, .xls e .xlsb. Máximo recomendado: 100 MB por arquivo.")
    if files:
        oversized = [file.name for file in files if len(file.getvalue()) > MAX_FILE_SIZE_MB * 1024 * 1024]
        if oversized:
            st.error("Arquivos acima do limite: " + ", ".join(oversized))
        elif st.button("Pré-validar arquivos", type="primary"):
            st.session_state.prevalidation = prevalidate(files, sheet_name, auto_header, header_line)
    reports = st.session_state.prevalidation
    if reports:
        st.markdown("#### Pré-validação")
        overview = pd.DataFrame([{key: value for key, value in report.items() if key not in {"values", "erro"}} for report in reports])
        st.dataframe(overview, use_container_width=True, hide_index=True)
        errors = [report for report in reports if report["status"] != "OK"]
        if errors:
            with st.expander("Ver orientações de correção", expanded=True):
                for report in errors:
                    st.error(f"{report['arquivo']}: {report['erro']}")
        st.markdown("#### Filtros de extração")
        f1, f2, f3 = st.columns(3)
        cfop = f1.multiselect("CFOP", available_values(reports, "cfop"))
        tes = f2.multiselect("TES", available_values(reports, "tes"))
        movimento = f3.multiselect("Tipo de Movimento", available_values(reports, "tipo_movimento"))
        f4, f5, f6 = st.columns(3)
        filial = f4.multiselect("filial", available_values(reports, "filial"))
        uf = f5.multiselect("UF", available_values(reports, "uf"))
        descricao = f6.text_input("Descrição do Produto")
        f7, f8 = st.columns(2)
        periodo_inicial = f7.date_input("Período inicial", value=None)
        periodo_final = f8.date_input("Período final", value=None)
        configuration = {
            "sheet_name": sheet_name, "auto_header": auto_header, "header_line": header_line,
            "mapping": parse_mapping(mapping_text),
            "filters": {"cfop": cfop, "tes": tes, "tipo_movimento": movimento, "filial": filial, "uf": uf, "descricao": descricao, "periodo_inicial": str(periodo_inicial) if periodo_inicial else "", "periodo_final": str(periodo_final) if periodo_final else ""},
        }
        if st.button("Consolidar e gerar relatórios", type="primary", use_container_width=True, disabled=not files):
            run_consolidation(user_id, files, configuration)
    if st.session_state.result_df is not None:
        st.markdown("#### Resultado consolidado")
        st.dataframe(st.session_state.result_df.head(500), use_container_width=True, hide_index=True)
        render_quality(st.session_state.quality, st.session_state.summary_df)
        result = st.session_state.result_df
        summary = st.session_state.summary_df
        col_excel, col_csv = st.columns(2)
        with col_excel:
            st.download_button("Baixar Excel consolidado", criar_excel_em_memoria(result, summary, sheet_name), f"consolidado_{datetime.now():%Y%m%d_%H%M%S}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        with col_csv:
            st.download_button("Baixar CSV brasileiro", result.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"), f"consolidado_{datetime.now():%Y%m%d_%H%M%S}.csv", "text/csv", use_container_width=True)


def run_consolidation(user_id: str, files: list[Any], configuration: dict[str, Any]) -> None:
    job = repository.create_job(user_id, f"Consolidação {datetime.now():%d/%m/%Y %H:%M}", configuration, len(files))
    frames, summary_rows = [], []
    try:
        progress = st.progress(0, text="Iniciando processamento...")
        for index, file in enumerate(files, start=1):
            progress.progress(index / len(files), text=f"Lendo {index}/{len(files)}: {file.name}")
            payload = file.getvalue()
            result = ler_planilha_robusta(io.BytesIO(payload), configuration["sheet_name"], configuration["auto_header"], None if configuration["auto_header"] else configuration["header_line"] - 1, True, True, None, file.name)
            source_total = 0.0
            filtered_rows = 0
            storage_path = repository.upload_bytes(user_id, job["id"], file.name, payload, file.type or "application/octet-stream")
            if result.status == "OK" and result.df is not None:
                df = result.df.rename(columns=configuration["mapping"])
                source_total = quality_summary(df)["total"]
                filtered = apply_filters(df, configuration["filters"])
                filtered_rows = len(filtered)
                frames.append(filtered)
                status = "processed"
            else:
                status = "failed"
            metadata = {"original_name": file.name, "storage_path": storage_path, "detected_sheet_name": result.aba, "header_row": (result.header_row_0based + 1 if result.header_row_0based is not None else None), "status": status, "read_rows": result.linhas, "filtered_rows": filtered_rows, "source_total": source_total, "error_message": result.erro}
            repository.add_processing_file(job["id"], metadata)
            summary_rows.append({"Arquivo": file.name, "Status": "OK" if result.status == "OK" else "FALHA", "Aba encontrada": result.aba or "-", "Linha de cabeçalho": metadata["header_row"], "Linhas lidas": result.linhas, "Linhas filtradas": filtered_rows, "Total de origem": source_total, "Orientação": result.erro or "Arquivo processado com sucesso."})
        progress.empty()
        consolidated = consolidar_planilhas(frames)
        summary = pd.DataFrame(summary_rows)
        quality = quality_summary(consolidated)
        excel = criar_excel_em_memoria(consolidated, summary, configuration["sheet_name"])
        csv = consolidated.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig")
        report = criar_excel_em_memoria(summary, summary, "Relatório")
        excel_path = repository.upload_bytes(user_id, job["id"], "consolidado.xlsx", excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        csv_path = repository.upload_bytes(user_id, job["id"], "consolidado.csv", csv, "text/csv")
        report_path = repository.upload_bytes(user_id, job["id"], "relatorio_processamento.xlsx", report, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        status = "completed" if all(row["Status"] == "OK" for row in summary_rows) else "completed_with_errors"
        repository.finish_job(job["id"], status=status, valid_files=sum(row["Status"] == "OK" for row in summary_rows), total_rows=sum(row["Linhas lidas"] for row in summary_rows), filtered_rows=len(consolidated), output_path=excel_path, csv_path=csv_path, report_path=report_path)
        repository.audit(user_id, "processing.completed", {"files": len(files), "rows": len(consolidated)}, job["id"])
        st.session_state.result_df, st.session_state.summary_df, st.session_state.quality, st.session_state.last_job = consolidated, summary, quality, job["id"]
        st.success(f"Consolidação concluída: {len(consolidated):,} linhas disponíveis para download.")
    except Exception as exc:
        repository.finish_job(job["id"], status="failed", valid_files=0, total_rows=0, filtered_rows=0, output_path=None, csv_path=None, report_path=None, error_message=str(exc))
        repository.audit(user_id, "processing.failed", {"error": str(exc)}, job["id"])
        st.error(f"Não foi possível concluir o processamento: {exc}")


def profiles_screen(user_id: str) -> None:
    st.title("Perfis")
    st.caption("Salve conjuntos reutilizáveis de aba, cabeçalho, filtros e mapeamentos.")
    with st.form("profile_save"):
        name = st.text_input("Nome do perfil", value="AGLO – Relatório Base")
        sheet = st.text_input("Aba do perfil", value=DEFAULT_SHEET_NAME)
        header = st.number_input("Linha de cabeçalho", min_value=1, value=2)
        mapping = st.text_area("Mapeamentos", value="CFOP Fiscal => CFOP\nC.F.O.P. => CFOP")
        shared = st.checkbox("Disponibilizar para outros usuários", value=False)
        if st.form_submit_button("Salvar perfil", type="primary"):
            try:
                repository.save_profile(user_id, name, {"sheet_name": sheet, "auto_header": True, "header_line": int(header), "mapping": parse_mapping(mapping), "filters": {}}, shared)
                st.success("Perfil salvo.")
            except Exception as exc:
                st.error(str(exc))
    profiles = repository.list_profiles(user_id)
    if profiles:
        st.dataframe(pd.DataFrame([{"Nome": item["name"], "Compartilhado": item["is_shared"], "Atualizado em": item["updated_at"]} for item in profiles]), use_container_width=True, hide_index=True)


def history_screen(user_id: str) -> None:
    st.title("Histórico")
    jobs = repository.list_jobs(user_id)
    if not jobs:
        st.info("Nenhum processamento foi registrado ainda.")
        return
    for job in jobs:
        with st.expander(f"{job['name']} · {job['status']} · {job['created_at']}"):
            a, b, c = st.columns(3)
            a.metric("Arquivos", job.get("total_files", 0))
            b.metric("Linhas lidas", job.get("total_rows", 0))
            c.metric("Linhas filtradas", job.get("filtered_rows", 0))
            config = job.get("configuration") or {}
            st.caption("Filtros aplicados: " + json.dumps(config.get("filters", {}), ensure_ascii=False))
            downloads = [("Excel consolidado", job.get("output_url")), ("CSV brasileiro", job.get("csv_url")), ("Relatório de processamento", job.get("report_url"))]
            for label, url in downloads:
                if url:
                    st.link_button(label, url)
            files = repository.list_job_files(job["id"])
            if files:
                st.dataframe(pd.DataFrame(files)[["original_name", "status", "detected_sheet_name", "read_rows", "filtered_rows", "source_total", "error_message"]], use_container_width=True, hide_index=True)


def admin_screen(user_id: str) -> None:
    st.title("Administração")
    st.caption("Gerencie os dois perfis de acesso permitidos: admin e usuário.")
    users = repository.list_users()
    if not users:
        st.info("Nenhum usuário cadastrado.")
        return
    for user in users:
        c1, c2, c3 = st.columns([3, 1, 1])
        c1.write(user.get("display_name") or user.get("email"))
        role = c2.selectbox("Perfil", ["admin", "user"], index=0 if user.get("role") == "admin" else 1, key=f"role_{user['id']}")
        if c3.button("Atualizar", key=f"save_{user['id']}"):
            try:
                repository.set_role(user_id, user["id"], role)
                st.success("Perfil atualizado.")
            except Exception as exc:
                st.error(str(exc))


def main() -> None:
    inject_style()
    init_state()
    if not st.session_state.auth:
        login_screen()
        return
    profile = repository.get_user_profile(st.session_state.auth["id"])
    if not profile:
        st.warning("Seu perfil está sendo preparado. Atualize a página em alguns segundos.")
        return
    current = sidebar(profile)
    user_id = st.session_state.auth["id"]
    if current == "Visão geral":
        dashboard(user_id)
    elif current == "Consolidar":
        consolidar_screen(user_id)
    elif current == "Perfis":
        profiles_screen(user_id)
    elif current in {"Histórico", "Relatórios"}:
        history_screen(user_id)
    elif current == "Administração" and profile.get("role") == "admin":
        admin_screen(user_id)


if __name__ == "__main__":
    main()
