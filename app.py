"""Consolidador de Apuração — acesso direto via Streamlit."""

from __future__ import annotations

import io
from datetime import datetime
from typing import Any, Iterable

import pandas as pd
import streamlit as st

from consolidate_relatorio_base import (
    DEFAULT_SHEET_NAME,
    MAX_FILE_SIZE_MB,
    SUPPORTED_EXCEL_EXTENSIONS,
    criar_excel_em_memoria,
    consolidar_planilhas,
    ler_planilha_robusta,
)


st.set_page_config(page_title="Consolidador de Apuração", page_icon="CA", layout="wide", initial_sidebar_state="expanded")


def inject_style() -> None:
    st.markdown(
        """
        <style>
        :root { --ink:#10172b; --line:#e7eaf1; --wash:#f6f7fb; --violet:#5b5ce2; }
        .stApp { background:var(--wash); color:var(--ink); }
        [data-testid="stSidebar"] { background:#fff; border-right:1px solid var(--line); }
        [data-testid="stSidebar"] > div:first-child { padding-top:1.2rem; }
        h1,h2,h3 { letter-spacing:-.035em; color:var(--ink); }
        .hero { background:linear-gradient(125deg,#070d20 0%,#111c3c 58%,#463ca5 160%); color:#fff; border-radius:24px; padding:2.6rem 2.8rem; margin:.2rem 0 1.7rem; box-shadow:0 18px 50px rgba(15,23,54,.17); }
        .hero .eyebrow { color:#d8d9ff; font-size:.78rem; font-weight:700; letter-spacing:.13em; text-transform:uppercase; }
        .hero h1 { color:#fff; font-size:2.45rem; margin:.5rem 0 .65rem; max-width:720px; line-height:1.08; }
        .hero p { color:#cbd5ef; font-size:1.03rem; margin:0; max-width:680px; }
        .metric-card { background:#fff; border:1px solid var(--line); border-radius:17px; padding:1.15rem 1.25rem; min-height:128px; }
        .metric-card .label { color:#7a8495; font-weight:700; font-size:.72rem; letter-spacing:.11em; text-transform:uppercase; }
        .metric-card .value { font-size:2.05rem; font-weight:750; letter-spacing:-.06em; margin:.45rem 0 .1rem; color:#17203b; }
        .metric-card .caption { color:#7a8495; font-size:.84rem; }
        .section-card { background:#fff; border:1px solid var(--line); border-radius:17px; padding:1.35rem; margin-bottom:1rem; }
        .section-title { font-size:1.1rem; font-weight:750; color:#1b2541; margin-bottom:.25rem; }
        .section-subtitle { color:#6d778a; font-size:.9rem; margin-bottom:1.1rem; }
        .stButton > button, .stDownloadButton > button { border-radius:10px; font-weight:650; min-height:2.65rem; }
        .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] { background:#5859dc; border-color:#5859dc; }
        .stDataFrame { border:1px solid var(--line); border-radius:12px; overflow:hidden; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def default_configuration() -> dict[str, Any]:
    return {
        "sheet_name": DEFAULT_SHEET_NAME,
        "auto_header": True,
        "header_line": 2,
        "mapping": {"CFOP Fiscal": "CFOP", "C.F.O.P.": "CFOP"},
        "filters": {},
    }


STANDARD_MAPPING_OPTIONS: dict[str, tuple[str, str]] = {
    "CFOP Fiscal → CFOP": ("CFOP Fiscal", "CFOP"),
    "C.F.O.P. → CFOP": ("C.F.O.P.", "CFOP"),
    "Tp. Mov → Tipo de Movimento": ("Tp. Mov", "Tipo de Movimento"),
    "Movimento → Tipo de Movimento": ("Movimento", "Tipo de Movimento"),
    "Desc. Produto → Descrição do Produto": ("Desc. Produto", "Descrição do Produto"),
    "Descricao Produto → Descrição do Produto": ("Descricao Produto", "Descrição do Produto"),
    "Loja → filial": ("Loja", "filial"),
    "Estado → UF": ("Estado", "UF"),
}


def init_state() -> None:
    defaults = {
        "prevalidation": [],
        "result_df": None,
        "summary_df": None,
        "quality": None,
        "session_jobs": [],
        "profiles": {"AGLO – Relatório Base": default_configuration()},
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
        return {"records": 0, "duplicates": 0, "empty": pd.DataFrame(), "cfop": pd.DataFrame(), "tes": pd.DataFrame(), "total": 0.0}
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
        return pd.DataFrame(columns=["Valor", "Registros"]) if not col else df[col].astype(str).str.strip().value_counts().head(12).rename_axis("Valor").reset_index(name="Registros")
    return {"records": len(df), "duplicates": duplicates, "empty": pd.DataFrame(empty), "cfop": distribution(candidates["CFOP"]), "tes": distribution(candidates["TES"]), "total": total}


def parse_mapping(text: str) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for row in text.splitlines():
        if "=>" in row:
            source, target = (part.strip() for part in row.split("=>", 1))
            if source and target:
                mapping[source] = target
    return mapping


def mapping_to_text(mapping: dict[str, str]) -> str:
    return "\n".join(f"{source} => {target}" for source, target in mapping.items())


def profile_mapping_editor(saved_mapping: dict[str, str]) -> dict[str, str]:
    selected_presets = [
        label
        for label, pair in STANDARD_MAPPING_OPTIONS.items()
        if saved_mapping.get(pair[0]) == pair[1]
    ]
    preset_pairs = set(STANDARD_MAPPING_OPTIONS.values())
    custom_rows = [
        {"Coluna de origem": source, "Nome padronizado": target}
        for source, target in saved_mapping.items()
        if (source, target) not in preset_pairs
    ]
    st.caption("Selecione equivalências frequentes ou acrescente regras específicas dos seus arquivos.")
    selected = st.multiselect(
        "Mapeamentos sugeridos",
        options=list(STANDARD_MAPPING_OPTIONS),
        default=selected_presets,
        help="Cada seleção converte a coluna de origem para o nome padronizado no relatório consolidado.",
    )
    edited_rows = st.data_editor(
        pd.DataFrame(custom_rows, columns=["Coluna de origem", "Nome padronizado"]),
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Coluna de origem": st.column_config.TextColumn("Coluna de origem", required=True),
            "Nome padronizado": st.column_config.TextColumn("Nome padronizado", required=True),
        },
    )
    mapping = {source: target for label in selected for source, target in [STANDARD_MAPPING_OPTIONS[label]]}
    for _, row in edited_rows.iterrows():
        source = clean_text(row.get("Coluna de origem"))
        target = clean_text(row.get("Nome padronizado"))
        if source and target:
            mapping[source] = target
    return mapping


def available_values(reports: list[dict[str, Any]], field: str) -> list[str]:
    return sorted({item for report in reports for item in report.get("values", {}).get(field, []) if item})


def apply_filters(df: pd.DataFrame, filters: dict[str, Any]) -> pd.DataFrame:
    work = df.copy()
    choices = {"cfop": ["CFOP", "C.F.O.P.", "CFOP Fiscal"], "tes": ["TES", "T.E.S"], "tipo_movimento": ["Tp. Mov", "Tipo de Movimento", "Movimento"], "filial": ["Filial", "Loja", "Estabelecimento"], "uf": ["UF", "Estado"]}
    for key, names in choices.items():
        selected, column = filters.get(key, []), find_column(work, names)
        if selected and column:
            work = work[work[column].map(canonical).isin({canonical(item) for item in selected})]
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


def sidebar() -> str:
    with st.sidebar:
        st.markdown("### Consolidador")
        st.caption("Apuração fiscal · acesso direto")
        current = st.radio("Navegação", ["Visão geral", "Consolidar", "Perfis", "Histórico da sessão"], label_visibility="collapsed")
        st.divider()
        st.caption("Os arquivos e relatórios ficam disponíveis somente durante esta sessão. Baixe os resultados antes de encerrar a página.")
    return current


def dashboard() -> None:
    jobs = st.session_state.session_jobs
    rows = sum(int(job["rows"]) for job in jobs)
    files = sum(int(job["files"]) for job in jobs)
    st.markdown("<div class='hero'><div class='eyebrow'>Central de operações</div><h1>Consolidação fiscal com clareza e rastreabilidade.</h1><p>Pré-valide arquivos, padronize colunas, aplique filtros e gere relatórios prontos para conferência.</p></div>", unsafe_allow_html=True)
    for column, label, value, caption in zip(st.columns(3), ["Processamentos", "Arquivos analisados", "Linhas consolidadas"], [len(jobs), files, rows], ["nesta sessão", "nesta sessão", "com auditoria"]):
        with column:
            st.markdown(f"<div class='metric-card'><div class='label'>{label}</div><div class='value'>{value:,}</div><div class='caption'>{caption}</div></div>", unsafe_allow_html=True)
    st.markdown("<div class='section-card'><div class='section-title'>Fluxo de trabalho</div><div class='section-subtitle'>1. Pré-validar arquivos · 2. Padronizar e filtrar · 3. Consolidar e exportar · 4. Conferir o histórico da sessão.</div></div>", unsafe_allow_html=True)


def prevalidate(files: list[Any], sheet_name: str, auto_header: bool, header_line: int) -> list[dict[str, Any]]:
    reports = []
    for file in files:
        result = ler_planilha_robusta(io.BytesIO(file.getvalue()), sheet_name, auto_header, None if auto_header else header_line - 1, True, False, None, file.name)
        values: dict[str, list[str]] = {"cfop": [], "tes": [], "tipo_movimento": [], "filial": [], "uf": []}
        if result.status == "OK" and result.df is not None:
            for key, names in {"cfop": ["CFOP", "C.F.O.P.", "CFOP Fiscal"], "tes": ["TES"], "tipo_movimento": ["Tp. Mov", "Tipo de Movimento"], "filial": ["Filial"], "uf": ["UF", "Estado"]}.items():
                column = find_column(result.df, names)
                if column:
                    values[key] = sorted(result.df[column].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())[:200]
        reports.append({"arquivo": file.name, "status": result.status, "aba": result.aba, "cabecalho": result.header_row_0based + 1 if result.header_row_0based is not None else None, "linhas": result.linhas, "colunas": result.colunas, "erro": result.erro, "values": values})
    return reports


def render_quality(quality: dict[str, Any], per_file: pd.DataFrame) -> None:
    st.markdown("<div class='section-card'><div class='section-title'>Qualidade dos dados</div><div class='section-subtitle'>Indicadores gerados sobre o resultado consolidado e os arquivos de origem.</div>", unsafe_allow_html=True)
    a, b, c = st.columns(3)
    a.metric("Registros", f"{quality['records']:,}")
    b.metric("Duplicidades", f"{quality['duplicates']:,}")
    c.metric("Total contábil", f"R$ {quality['total']:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    left, right = st.columns(2)
    left.dataframe(quality["cfop"], use_container_width=True, hide_index=True)
    right.dataframe(quality["tes"], use_container_width=True, hide_index=True)
    st.caption("Campos para revisão")
    st.dataframe(quality["empty"], use_container_width=True, hide_index=True)
    st.caption("Conferência por arquivo")
    st.dataframe(per_file, use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)


def consolidation_screen() -> None:
    st.title("Consolidar")
    st.caption("Envie arquivos, confirme a estrutura identificada e produza relatórios fiscalmente rastreáveis.")
    names = list(st.session_state.profiles)
    selected = st.selectbox("Perfil de configuração", names, help="Perfis são guardados somente durante a sessão atual.")
    profile = st.session_state.profiles[selected]
    with st.expander("Configuração da leitura", expanded=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        sheet_name = c1.text_input("Nome da aba", value=profile["sheet_name"])
        auto_header = c2.checkbox("Detectar cabeçalho", value=profile["auto_header"])
        header_line = int(c3.number_input("Linha manual", min_value=1, value=int(profile["header_line"]), disabled=auto_header))
        mapping_text = st.text_area("Mapeamento de colunas", value=mapping_to_text(profile.get("mapping", {})), help="Uma equivalência por linha, no formato coluna de origem => nome padronizado.")
    files = st.file_uploader("Arquivos Excel", type=[item.lstrip(".") for item in SUPPORTED_EXCEL_EXTENSIONS], accept_multiple_files=True, help="Formatos aceitos: .xlsx, .xlsm, .xltx, .xltm, .xls e .xlsb. Máximo recomendado: 100 MB por arquivo.")
    if files:
        oversized = [file.name for file in files if len(file.getvalue()) > MAX_FILE_SIZE_MB * 1024 * 1024]
        if oversized:
            st.error("Arquivos acima do limite: " + ", ".join(oversized))
        elif st.button("Pré-validar arquivos", type="primary"):
            st.session_state.prevalidation = prevalidate(files, sheet_name, auto_header, header_line)
    reports = st.session_state.prevalidation
    if not reports:
        return
    st.markdown("#### Pré-validação")
    st.dataframe(pd.DataFrame([{key: value for key, value in report.items() if key not in {"values", "erro"}} for report in reports]), use_container_width=True, hide_index=True)
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
    inicio = f7.date_input("Período inicial", value=None)
    fim = f8.date_input("Período final", value=None)
    configuration = {"sheet_name": sheet_name, "auto_header": auto_header, "header_line": header_line, "mapping": parse_mapping(mapping_text), "filters": {"cfop": cfop, "tes": tes, "tipo_movimento": movimento, "filial": filial, "uf": uf, "descricao": descricao, "periodo_inicial": str(inicio) if inicio else "", "periodo_final": str(fim) if fim else ""}}
    if st.button("Consolidar e gerar relatórios", type="primary", use_container_width=True, disabled=not files):
        run_consolidation(files, configuration)
    if st.session_state.result_df is not None:
        st.markdown("#### Resultado consolidado")
        st.dataframe(st.session_state.result_df.head(500), use_container_width=True, hide_index=True)
        render_quality(st.session_state.quality, st.session_state.summary_df)
        excel = criar_excel_em_memoria(st.session_state.result_df, st.session_state.summary_df, sheet_name)
        csv = st.session_state.result_df.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
        a, b = st.columns(2)
        a.download_button("Baixar Excel consolidado", excel, f"consolidado_{datetime.now():%Y%m%d_%H%M%S}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        b.download_button("Baixar CSV brasileiro", csv, f"consolidado_{datetime.now():%Y%m%d_%H%M%S}.csv", "text/csv", use_container_width=True)


def run_consolidation(files: list[Any], configuration: dict[str, Any]) -> None:
    frames, rows = [], []
    progress = st.progress(0, text="Iniciando processamento...")
    for index, file in enumerate(files, start=1):
        progress.progress(index / len(files), text=f"Lendo {index}/{len(files)}: {file.name}")
        result = ler_planilha_robusta(io.BytesIO(file.getvalue()), configuration["sheet_name"], configuration["auto_header"], None if configuration["auto_header"] else configuration["header_line"] - 1, True, True, None, file.name)
        source_total, filtered_rows = 0.0, 0
        if result.status == "OK" and result.df is not None:
            df = result.df.rename(columns=configuration["mapping"])
            source_total = quality_summary(df)["total"]
            filtered = apply_filters(df, configuration["filters"])
            filtered_rows = len(filtered)
            frames.append(filtered)
        rows.append({"Arquivo": file.name, "Status": "OK" if result.status == "OK" else "FALHA", "Aba encontrada": result.aba or "-", "Linha de cabeçalho": result.header_row_0based + 1 if result.header_row_0based is not None else None, "Linhas lidas": result.linhas, "Linhas filtradas": filtered_rows, "Total de origem": source_total, "Orientação": result.erro or "Arquivo processado com sucesso."})
    progress.empty()
    consolidated, summary = consolidar_planilhas(frames), pd.DataFrame(rows)
    st.session_state.result_df, st.session_state.summary_df, st.session_state.quality = consolidated, summary, quality_summary(consolidated)
    st.session_state.session_jobs.insert(0, {"name": f"Consolidação {datetime.now():%d/%m/%Y %H:%M}", "files": len(files), "rows": len(consolidated), "summary": summary, "excel": criar_excel_em_memoria(consolidated, summary, configuration["sheet_name"]), "csv": consolidated.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")})
    st.success(f"Consolidação concluída: {len(consolidated):,} linhas disponíveis para download.")


def profiles_screen() -> None:
    st.title("Perfis")
    st.caption("Crie perfis com regras de leitura e mapeamentos por seleção. Eles ficam disponíveis enquanto esta sessão permanecer aberta.")
    profile_names = list(st.session_state.profiles)
    base_name = st.selectbox("Editar a partir do perfil", profile_names, help="Use um perfil existente como ponto de partida.")
    base_profile = st.session_state.profiles[base_name]
    with st.form("profile_save"):
        name = st.text_input("Nome do perfil", value=base_name)
        sheet = st.text_input("Aba do perfil", value=base_profile["sheet_name"])
        header = st.number_input("Linha de cabeçalho", min_value=1, value=int(base_profile["header_line"]))
        st.markdown("#### Mapeamentos de colunas")
        mapping = profile_mapping_editor(base_profile.get("mapping", {}))
        if st.form_submit_button("Salvar perfil", type="primary"):
            if not clean_text(name):
                st.error("Informe um nome para o perfil.")
            else:
                st.session_state.profiles[name.strip()] = {"sheet_name": sheet, "auto_header": True, "header_line": int(header), "mapping": mapping, "filters": {}}
                st.success("Perfil salvo para a sessão atual.")
    overview = pd.DataFrame(
        [
            {"Perfil": profile_name, "Mapeamentos": len(config.get("mapping", {})), "Aba": config.get("sheet_name", "-")}
            for profile_name, config in st.session_state.profiles.items()
        ]
    )
    st.dataframe(overview, use_container_width=True, hide_index=True)


def history_screen() -> None:
    st.title("Histórico da sessão")
    jobs = st.session_state.session_jobs
    if not jobs:
        st.info("Nenhum processamento foi registrado nesta sessão.")
        return
    for index, job in enumerate(jobs):
        with st.expander(f"{job['name']} · {job['files']} arquivo(s) · {job['rows']:,} linhas"):
            st.dataframe(job["summary"], use_container_width=True, hide_index=True)
            a, b = st.columns(2)
            a.download_button("Baixar Excel", job["excel"], f"consolidado_sessao_{index + 1}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"session_excel_{index}")
            b.download_button("Baixar CSV brasileiro", job["csv"], f"consolidado_sessao_{index + 1}.csv", "text/csv", key=f"session_csv_{index}")


def main() -> None:
    inject_style()
    init_state()
    current = sidebar()
    if current == "Visão geral":
        dashboard()
    elif current == "Consolidar":
        consolidation_screen()
    elif current == "Perfis":
        profiles_screen()
    else:
        history_screen()


if __name__ == "__main__":
    main()
