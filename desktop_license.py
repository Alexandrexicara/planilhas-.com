"""
Sistema de licenças desktop com proteção anti-pirataria.

Estratégia:
- 1 licença por organização (organization_id UNIQUE)
- Token único gerado pelo servidor (não regravável)
- Hardware fingerprint amarrado na primeira ativação
- Servidor recusa se outro PC tentar usar o mesmo token
- Admin pode resetar hardware via painel (libera novo PC)

Suporta PostgreSQL (Render) e SQLite (fallback local).
"""
import os
import secrets
from datetime import datetime


def _get_db_url():
    return os.environ.get('DATABASE_URL')


def _connect():
    """Conecta ao PostgreSQL (Render) ou SQLite (local)."""
    db_url = _get_db_url()
    if db_url:
        import psycopg2
        conn = psycopg2.connect(db_url)
        conn.autocommit = True
        return conn, True
    import sqlite3
    try:
        from planilhas_paths import ensure_from_resource
        db_path = ensure_from_resource("acesso_web.db")
    except Exception:
        db_path = "acesso_web.db"
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn, False


def init_license_table():
    """Cria tabela desktop_licenses se não existir."""
    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS desktop_licenses (
                        id SERIAL PRIMARY KEY,
                        organization_id INTEGER UNIQUE NOT NULL,
                        token VARCHAR(64) UNIQUE NOT NULL,
                        hardware_id VARCHAR(255),
                        activated_at TIMESTAMP,
                        last_check TIMESTAMP,
                        revoked INTEGER DEFAULT 0,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS desktop_licenses (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    organization_id INTEGER UNIQUE NOT NULL,
                    token TEXT UNIQUE NOT NULL,
                    hardware_id TEXT,
                    activated_at TEXT,
                    last_check TEXT,
                    revoked INTEGER DEFAULT 0,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()
    finally:
        conn.close()


def _row_to_dict(row, cols):
    if row is None:
        return None
    if isinstance(row, dict):
        return row
    return {c: row[i] for i, c in enumerate(cols)}


_COLS = ["id", "organization_id", "token", "hardware_id",
         "activated_at", "last_check", "revoked", "created_at"]


def get_license_by_org(organization_id):
    """Retorna a licença da organização ou None."""
    if not organization_id:
        return None
    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id, organization_id, token, hardware_id, "
                    "activated_at, last_check, revoked, created_at "
                    "FROM desktop_licenses WHERE organization_id = %s",
                    (organization_id,)
                )
                row = cur.fetchone()
        else:
            cur = conn.execute(
                "SELECT id, organization_id, token, hardware_id, "
                "activated_at, last_check, revoked, created_at "
                "FROM desktop_licenses WHERE organization_id = ?",
                (organization_id,)
            )
            row = cur.fetchone()
        return _row_to_dict(row, _COLS)
    finally:
        conn.close()


def get_license_by_token(token):
    """Retorna a licença pelo token ou None."""
    if not token:
        return None
    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT id, organization_id, token, hardware_id, "
                    "activated_at, last_check, revoked, created_at "
                    "FROM desktop_licenses WHERE token = %s",
                    (token,)
                )
                row = cur.fetchone()
        else:
            cur = conn.execute(
                "SELECT id, organization_id, token, hardware_id, "
                "activated_at, last_check, revoked, created_at "
                "FROM desktop_licenses WHERE token = ?",
                (token,)
            )
            row = cur.fetchone()
        return _row_to_dict(row, _COLS)
    finally:
        conn.close()


def create_or_get_license(organization_id):
    """Gera (ou recupera) a única licença desktop da organização."""
    if not organization_id:
        raise ValueError("organization_id obrigatório")

    existing = get_license_by_org(organization_id)
    if existing:
        return existing

    # Token único e seguro
    token = secrets.token_urlsafe(32)

    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute(
                    "INSERT INTO desktop_licenses (organization_id, token) "
                    "VALUES (%s, %s)",
                    (organization_id, token)
                )
        else:
            conn.execute(
                "INSERT INTO desktop_licenses (organization_id, token) "
                "VALUES (?, ?)",
                (organization_id, token)
            )
            conn.commit()
    finally:
        conn.close()

    return get_license_by_org(organization_id)


def activate_license(token, hardware_id):
    """Ativa a licença amarrando ao hardware_id (primeira ativação).

    Retorna dict: {ok: bool, motivo: str}
    """
    if not token or not hardware_id:
        return {"ok": False, "motivo": "Token ou hardware_id ausente"}

    lic = get_license_by_token(token)
    if not lic:
        return {"ok": False, "motivo": "Token inválido"}
    if lic.get("revoked"):
        return {"ok": False, "motivo": "Licença revogada"}

    existing_hw = lic.get("hardware_id")
    if existing_hw and existing_hw != hardware_id:
        return {
            "ok": False,
            "motivo": ("Esta licença já está ativada em outro computador. "
                       "Peça ao administrador para resetar a licença.")
        }

    if existing_hw == hardware_id:
        # Já ativado nesse mesmo PC -> renova last_check
        _update_last_check(lic["id"])
        return {"ok": True, "motivo": "Licença já ativa neste computador",
                "organization_id": lic["organization_id"]}

    # Primeira ativação
    now = datetime.utcnow().isoformat()
    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE desktop_licenses "
                    "SET hardware_id=%s, activated_at=%s, last_check=%s "
                    "WHERE id=%s",
                    (hardware_id, now, now, lic["id"])
                )
        else:
            conn.execute(
                "UPDATE desktop_licenses "
                "SET hardware_id=?, activated_at=?, last_check=? "
                "WHERE id=?",
                (hardware_id, now, now, lic["id"])
            )
            conn.commit()
    finally:
        conn.close()

    return {"ok": True, "motivo": "Licença ativada com sucesso",
            "organization_id": lic["organization_id"]}


def verify_license(token, hardware_id):
    """Verifica se token+hardware_id estão válidos.

    Retorna dict: {ok: bool, motivo: str}
    """
    if not token or not hardware_id:
        return {"ok": False, "motivo": "Token ou hardware_id ausente"}

    lic = get_license_by_token(token)
    if not lic:
        return {"ok": False, "motivo": "Token inválido"}
    if lic.get("revoked"):
        return {"ok": False, "motivo": "Licença revogada"}
    if not lic.get("hardware_id"):
        return {"ok": False, "motivo": "Licença ainda não ativada"}
    if lic["hardware_id"] != hardware_id:
        return {"ok": False, "motivo": "Hardware diferente do registrado"}

    _update_last_check(lic["id"])
    return {"ok": True, "motivo": "Licença válida",
            "organization_id": lic["organization_id"]}


def _update_last_check(license_id):
    now = datetime.utcnow().isoformat()
    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE desktop_licenses SET last_check=%s WHERE id=%s",
                    (now, license_id)
                )
        else:
            conn.execute(
                "UPDATE desktop_licenses SET last_check=? WHERE id=?",
                (now, license_id)
            )
            conn.commit()
    finally:
        conn.close()


def reset_license(organization_id):
    """Limpa hardware_id (libera nova ativação). Mantém o mesmo token."""
    if not organization_id:
        return False
    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE desktop_licenses "
                    "SET hardware_id=NULL, activated_at=NULL, last_check=NULL "
                    "WHERE organization_id=%s",
                    (organization_id,)
                )
        else:
            conn.execute(
                "UPDATE desktop_licenses "
                "SET hardware_id=NULL, activated_at=NULL, last_check=NULL "
                "WHERE organization_id=?",
                (organization_id,)
            )
            conn.commit()
    finally:
        conn.close()
    return True


def revoke_license(organization_id):
    """Revoga a licença (cancelamento)."""
    conn, is_pg = _connect()
    try:
        if is_pg:
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE desktop_licenses SET revoked=1 "
                    "WHERE organization_id=%s",
                    (organization_id,)
                )
        else:
            conn.execute(
                "UPDATE desktop_licenses SET revoked=1 "
                "WHERE organization_id=?",
                (organization_id,)
            )
            conn.commit()
    finally:
        conn.close()
    return True


# ============================================================
# Cliente (rodando dentro do .exe) - hardware fingerprint
# ============================================================

def gerar_hardware_id():
    """Gera ID único do hardware atual (CPU + MAC + hostname + disco).

    Resultado é estável no mesmo PC e diferente entre PCs.
    Não requer privilégios de administrador.
    """
    import hashlib
    import platform
    import uuid

    partes = []
    try:
        partes.append(str(uuid.getnode()))  # MAC address (numérico)
    except Exception:
        partes.append("nomac")
    try:
        partes.append(platform.node() or "noname")  # Hostname
    except Exception:
        partes.append("noname")
    try:
        partes.append(platform.machine() or "noarch")  # x86_64 etc.
    except Exception:
        partes.append("noarch")
    try:
        partes.append(platform.processor() or "nocpu")
    except Exception:
        partes.append("nocpu")

    # No Windows tenta serial do disco C: (não exige admin)
    try:
        if os.name == 'nt':
            import subprocess
            r = subprocess.run(
                ["vol", "C:"], capture_output=True, text=True,
                timeout=3, shell=True
            )
            partes.append(r.stdout.strip())
    except Exception:
        pass

    raw = "|".join(partes)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()
