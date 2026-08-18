"""Persistência e autenticação para a versão Streamlit do Consolidador."""

from __future__ import annotations

import json
import os
import re
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, Optional

import streamlit as st
from supabase import Client, create_client


def _secret(name: str) -> Optional[str]:
    try:
        value = st.secrets.get(name) or st.secrets.get("supabase", {}).get(name.lower())
    except Exception:
        value = os.getenv(name)
    return str(value).strip() if value else None


def _validate_connection_value(name: str, value: str) -> None:
    try:
        value.encode("ascii")
    except UnicodeEncodeError as exc:
        raise RuntimeError(
            f"O secret {name} contém um caractere não permitido. Copie apenas o valor real da chave no Supabase, "
            "sem setas, explicações, aspas extras ou texto de exemplo."
        ) from exc


@st.cache_resource
def get_supabase() -> Client:
    url = _secret("SUPABASE_URL")
    key = _secret("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        raise RuntimeError(
            "Supabase não configurado. Informe SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY nos secrets do Streamlit Cloud."
        )
    _validate_connection_value("SUPABASE_URL", url)
    _validate_connection_value("SUPABASE_SERVICE_ROLE_KEY", key)
    return create_client(url, key)


def configured() -> bool:
    return bool(_secret("SUPABASE_URL") and _secret("SUPABASE_SERVICE_ROLE_KEY"))


def sign_in(email: str, password: str) -> Dict[str, Any]:
    response = get_supabase().auth.sign_in_with_password({"email": email, "password": password})
    if not response.user or not response.session:
        raise RuntimeError("Não foi possível autenticar. Confirme seu e-mail e senha.")
    return {"id": response.user.id, "email": response.user.email or email, "access_token": response.session.access_token}


def sign_up(email: str, password: str, display_name: str) -> None:
    get_supabase().auth.sign_up({"email": email, "password": password, "options": {"data": {"name": display_name}}})


def get_user_profile(user_id: str) -> Optional[Dict[str, Any]]:
    response = get_supabase().table("app_users").select("*").eq("id", user_id).limit(1).execute()
    return response.data[0] if response.data else None


def _safe_name(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]", "_", value)[-140:]


def upload_bytes(user_id: str, job_id: str, name: str, payload: bytes, content_type: str) -> str:
    path = f"{user_id}/{job_id}/{uuid.uuid4().hex}_{_safe_name(name)}"
    get_supabase().storage.from_("fiscal-files").upload(path, payload, {"content-type": content_type, "upsert": "false"})
    return path


def signed_url(path: Optional[str], expires_in: int = 3600) -> Optional[str]:
    if not path:
        return None
    response = get_supabase().storage.from_("fiscal-files").create_signed_url(path, expires_in)
    return response.get("signedURL") or response.get("signedUrl")


def create_job(user_id: str, name: str, configuration: Dict[str, Any], total_files: int) -> Dict[str, Any]:
    response = get_supabase().table("processing_jobs").insert({
        "user_id": user_id,
        "name": name,
        "configuration": configuration,
        "total_files": total_files,
        "status": "processing",
    }).execute()
    return response.data[0]


def finish_job(job_id: str, *, status: str, valid_files: int, total_rows: int, filtered_rows: int,
               output_path: Optional[str], csv_path: Optional[str], report_path: Optional[str], error_message: Optional[str] = None) -> None:
    get_supabase().table("processing_jobs").update({
        "status": status,
        "valid_files": valid_files,
        "total_rows": total_rows,
        "filtered_rows": filtered_rows,
        "output_path": output_path,
        "csv_path": csv_path,
        "report_path": report_path,
        "error_message": error_message,
        "completed_at": datetime.now(timezone.utc).isoformat(),
    }).eq("id", job_id).execute()


def add_processing_file(job_id: str, metadata: Dict[str, Any]) -> None:
    get_supabase().table("processing_files").insert({"job_id": job_id, **metadata}).execute()


def audit(user_id: str, action: str, details: Dict[str, Any], job_id: Optional[str] = None) -> None:
    get_supabase().table("audit_events").insert({"user_id": user_id, "job_id": job_id, "action": action, "details": details}).execute()


def save_profile(user_id: str, name: str, configuration: Dict[str, Any], is_shared: bool = False) -> None:
    get_supabase().table("configuration_profiles").upsert({
        "owner_id": user_id, "name": name, "configuration": configuration, "is_shared": is_shared,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }, on_conflict="owner_id,name").execute()
    audit(user_id, "profile.saved", {"name": name})


def list_profiles(user_id: str) -> list[Dict[str, Any]]:
    response = get_supabase().table("configuration_profiles").select("*").or_(f"owner_id.eq.{user_id},is_shared.eq.true").order("updated_at", desc=True).execute()
    return response.data or []


def list_jobs(user_id: str) -> list[Dict[str, Any]]:
    response = get_supabase().table("processing_jobs").select("*").eq("user_id", user_id).order("created_at", desc=True).limit(50).execute()
    jobs = response.data or []
    for job in jobs:
        job["output_url"] = signed_url(job.get("output_path"))
        job["csv_url"] = signed_url(job.get("csv_path"))
        job["report_url"] = signed_url(job.get("report_path"))
    return jobs


def list_job_files(job_id: str) -> list[Dict[str, Any]]:
    response = get_supabase().table("processing_files").select("*").eq("job_id", job_id).order("created_at").execute()
    return response.data or []


def list_users() -> list[Dict[str, Any]]:
    response = get_supabase().table("app_users").select("*").order("created_at").execute()
    return response.data or []


def set_role(actor_id: str, user_id: str, role: str) -> None:
    if role not in {"admin", "user"}:
        raise ValueError("Perfil inválido.")
    get_supabase().table("app_users").update({"role": role, "updated_at": datetime.now(timezone.utc).isoformat()}).eq("id", user_id).execute()
    audit(actor_id, "user.role_changed", {"user_id": user_id, "role": role})
