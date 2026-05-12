"""Módulo de acesso ao banco de dados para o sistema Planilhas.

Fornece funções para autenticação, gerenciamento de usuários,
organizações, convites e upload de imagens.
"""
import os
import sqlite3
import secrets
from datetime import datetime

from werkzeug.security import generate_password_hash, check_password_hash

from planilhas_paths import ensure_from_resource


DB_FILENAME = "acesso_web.db"


def _utcnow_str():
    """Retorna a data/hora atual em UTC como string formatada."""
    return datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")


def get_db_path():
    """Retorna o caminho completo do arquivo de banco de dados."""
    # No Render, usar diretório temporário
    if os.environ.get('RENDER'):
        db_dir = os.path.join('/tmp', 'planilhas')
        os.makedirs(db_dir, exist_ok=True)
        return os.path.join(db_dir, DB_FILENAME)

    return ensure_from_resource(DB_FILENAME)


def connect():
    """Conecta ao banco de dados SQLite e garante que as tabelas existam."""
    # Garantir que o banco existe no Render
    db_path = get_db_path()
    if os.environ.get('RENDER') and not os.path.exists(db_path):
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        conn = sqlite3.connect(db_path)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                organization_id INTEGER,
                nome TEXT NOT NULL,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'collab',
                ativo INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            )
        """)
    elif os.environ.get('RENDER') and os.path.exists(db_path):
        # Migração: verificar se a coluna senha existe e renomear para password_hash
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(users)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'senha' in columns and 'password_hash' not in columns:
            print("[MIGRAÇÃO] Renomeando coluna 'senha' para 'password_hash'")
            # SQLite não suporta ALTER COLUMN diretamente, precisa recriar a tabela
            cursor.execute("""
                CREATE TABLE users_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    organization_id INTEGER,
                    nome TEXT NOT NULL,
                    email TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL DEFAULT 'collab',
                    ativo INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL
                )
            """)
            cursor.execute("""
                INSERT INTO users_new (id, organization_id, nome, email, password_hash, role, ativo, created_at)
                SELECT id, organization_id, nome, email, senha, role, ativo, created_at FROM users
            """)
            cursor.execute("DROP TABLE users")
            cursor.execute("ALTER TABLE users_new RENAME TO users")
            conn.commit()
            print("[MIGRAÇÃO] Concluída com sucesso")
        conn.close()
    
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn
        conn.execute("""
            CREATE TABLE IF NOT EXISTS imagens_upload (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                nome_original TEXT NOT NULL,
                url TEXT NOT NULL,
                tipo TEXT,
                tamanho INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
        conn.close()

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Inicializa o banco de dados criando todas as tabelas necessárias."""
    conn = connect()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS organizations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            created_at TEXT NOT NULL,

            payment_status TEXT NOT NULL DEFAULT 'unpaid', -- unpaid | pending | paid
            payment_amount REAL NOT NULL DEFAULT 50.00,
            payment_txid TEXT,
            payment_qr_base64 TEXT,
            payment_pix_key TEXT,
            payment_updated_at TEXT
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            organization_id INTEGER,
            nome TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'collab', -- superadmin | owner | admin | collab
            ativo INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            FOREIGN KEY (organization_id) REFERENCES organizations (id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS invites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            organization_id INTEGER NOT NULL,
            code TEXT UNIQUE NOT NULL,
            role TEXT NOT NULL DEFAULT 'collab',
            email TEXT,
            created_at TEXT NOT NULL,
            used_at TEXT,
            used_by_user_id INTEGER,
            FOREIGN KEY (organization_id) REFERENCES organizations (id),
            FOREIGN KEY (used_by_user_id) REFERENCES users (id)
        )
        """
    )

    # Tabela para arquivamento de imagens
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS imagens_upload (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            nome_original TEXT NOT NULL,
            url TEXT NOT NULL,
            tipo TEXT,
            tamanho INTEGER,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    conn.commit()
    conn.close()


def any_organization_exists():
    """Verifica se existe pelo menos uma organização no banco."""
    conn = connect()
    try:
        row = conn.execute("SELECT 1 FROM organizations LIMIT 1").fetchone()
        return bool(row)
    finally:
        conn.close()


def create_organization(nome, payment_amount=50.00):
    """Cria uma nova organização no banco de dados."""
    conn = connect()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO organizations (nome, created_at, payment_status, payment_amount)
            VALUES (?, ?, 'unpaid', ?)
            """,
            (nome.strip(), _utcnow_str(), float(payment_amount)),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def create_user(organization_id, nome, email, senha, role="collab", ativo=1):
    """Cria um novo usuário vinculado a uma organização."""
    conn = connect()
    try:
        password_hash = generate_password_hash(senha)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO users (organization_id, nome, email, password_hash, role, ativo, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                organization_id,
                nome.strip(),
                email.strip().lower(),
                password_hash,
                role,
                int(ativo),
                _utcnow_str(),
            ),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def ensure_superadmin(email, senha):
    """Garante que existe um superadmin no sistema, criando se necessário."""
    if not email or not senha:
        return None

    email_norm = email.strip().lower()
    conn = connect()
    try:
        existing = conn.execute("SELECT id FROM users WHERE email = ?", (email_norm,)).fetchone()
        password_hash = generate_password_hash(senha)
        if existing:
            conn.execute(
                """
                UPDATE users
                SET password_hash = ?, role = 'superadmin', ativo = 1
                WHERE id = ?
                """,
                (password_hash, existing["id"]),
            )
            conn.commit()
            return existing["id"]

        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO users (organization_id, nome, email, password_hash, role, ativo, created_at)
            VALUES (NULL, 'Criador', ?, ?, 'superadmin', 1, ?)
            """,
            (email_norm, password_hash, _utcnow_str()),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def authenticate(email, senha):
    """Autentica um usuário pelo email e senha."""
    print("=== DEBUG AUTHENTICATE ===")
    print(f"Email: {email}")
    print(f"Senha: {senha}")

    if not email or not senha:
        print("Email ou senha vazios")
        return None

    conn = connect()
    try:
        print("Buscando usuário no banco...")
        row = conn.execute(
            """
            SELECT id, organization_id, nome, email, password_hash, role, ativo
            FROM users
            WHERE email = ?
            """,
            (email.strip().lower(),),
        ).fetchone()

        print(f"Usuário encontrado: {row}")

        if not row:
            print("Usuário não encontrado")
            return None

        if int(row["ativo"]) != 1:
            print("Usuário inativo")
            return None

        print(f"Verificando senha: {row['password_hash']} vs {senha}")
        if not check_password_hash(row["password_hash"], senha):
            print("Senha incorreta")
            return None

        print("Autenticação bem-sucedida!")
        return {
            "id": row["id"],
            "organization_id": row["organization_id"],
            "nome": row["nome"],
            "email": row["email"],
            "role": row["role"],
        }
    finally:
        conn.close()


def get_user(user_id):
    """Busca um usuário pelo ID."""
    conn = connect()
    try:
        row = conn.execute(
            """
            SELECT id, organization_id, nome, email, role, ativo
            FROM users
            WHERE id = ?
            """,
            (int(user_id),),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_organization(org_id):
    """Busca uma organização pelo ID."""
    if org_id is None:
        return None

    conn = connect()
    try:
        row = conn.execute(
            """
            SELECT id, nome, payment_status, payment_amount, payment_txid,
                   payment_qr_base64, payment_pix_key, payment_updated_at
            FROM organizations
            WHERE id = ?
            """,
            (int(org_id),),
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def organization_has_access(org_id):
    """Verifica se a organização tem acesso pago ao sistema."""
    org = get_organization(org_id)
    if not org:
        return False
    return org.get("payment_status") == "paid"


def set_organization_payment_pending(org_id, txid, qr_base64=None, pix_key=None):
    """Define o status de pagamento da organização como pendente."""
    conn = connect()
    try:
        conn.execute(
            """
            UPDATE organizations
            SET payment_status = 'pending',
                payment_txid = ?,
                payment_qr_base64 = ?,
                payment_pix_key = ?,
                payment_updated_at = ?
            WHERE id = ?
            """,
            (txid, qr_base64, pix_key, _utcnow_str(), int(org_id)),
        )
        conn.commit()
    finally:
        conn.close()


def set_organization_paid(org_id):
    """Define o status de pagamento da organização como pago."""
    conn = connect()
    try:
        conn.execute(
            """
            UPDATE organizations
            SET payment_status = 'paid',
                payment_updated_at = ?
            WHERE id = ?
            """,
            (_utcnow_str(), int(org_id)),
        )
        conn.commit()
    finally:
        conn.close()


def create_invite(organization_id, role="collab", email=None):
    """Cria um código de convite para uma organização."""
    code = secrets.token_urlsafe(10).replace("-", "").replace("_", "")
    conn = connect()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO invites (organization_id, code, role, email, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                int(organization_id),
                code,
                role,
                email.strip().lower() if email else None,
                _utcnow_str(),
            ),
        )
        conn.commit()
        return code
    finally:
        conn.close()


def redeem_invite(code, nome, email, senha):
    """Resgata um convite criando um novo usuário na organização."""
    if not code:
        return False, "Código de convite obrigatório", None

    code_norm = code.strip()
    email_norm = (email or "").strip().lower()
    conn = connect()
    try:
        invite = conn.execute(
            """
            SELECT id, organization_id, role, email, used_at
            FROM invites
            WHERE code = ?
            """,
            (code_norm,),
        ).fetchone()

        if not invite:
            return False, "Convite inválido", None
        if invite["used_at"]:
            return False, "Convite já utilizado", None
        if invite["email"] and invite["email"] != email_norm:
            return False, "Este convite é para outro email", None

        user_id = create_user(
            invite["organization_id"],
            nome,
            email_norm,
            senha,
            role=invite["role"]
        )
        conn.execute(
            """
            UPDATE invites
            SET used_at = ?, used_by_user_id = ?
            WHERE id = ?
            """,
            (_utcnow_str(), int(user_id), int(invite["id"])),
        )
        conn.commit()
        return True, "Cadastro realizado", user_id
    except sqlite3.IntegrityError:
        return False, "Este email já está cadastrado", None
    finally:
        conn.close()


def list_invites(organization_id, limit=50):
    """Lista convites ativos de uma organização."""
    conn = connect()
    try:
        rows = conn.execute(
            """
            SELECT id, code, role, email, created_at, used_at, used_by_user_id
            FROM invites
            WHERE organization_id = ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (int(organization_id), int(limit)),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def salvar_imagem(filename, nome_original, url, tipo=None, tamanho=None):
    """Salva registro de imagem enviada no banco."""
    conn = connect()
    try:
        conn.execute(
            """
            INSERT INTO imagens_upload (filename, nome_original, url, tipo, tamanho, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (filename, nome_original, url, tipo, tamanho, _utcnow_str()),
        )
        conn.commit()
        return True
    except sqlite3.Error as e:
        print(f"❌ Erro ao salvar imagem no banco: {e}")
        return False
    finally:
        conn.close()


def listar_imagens(limit=100, offset=0):
    """Lista imagens arquivadas ordenadas por data (mais recentes primeiro)."""
    conn = connect()
    try:
        rows = conn.execute(
            """
            SELECT id, filename, nome_original, url, tipo, tamanho, created_at
            FROM imagens_upload
            ORDER BY created_at DESC
            LIMIT ? OFFSET ?
            """,
            (int(limit), int(offset)),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def buscar_imagens(termo, limit=50):
    """Busca imagens por nome original."""
    conn = connect()
    try:
        rows = conn.execute(
            """
            SELECT id, filename, nome_original, url, tipo, tamanho, created_at
            FROM imagens_upload
            WHERE nome_original LIKE ?
            ORDER BY created_at DESC
            LIMIT ?
            """,
            (f"%{termo}%", int(limit)),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def list_users(organization_id, limit=200):
    """Lista usuários de uma organização."""
    conn = connect()
    try:
        rows = conn.execute(
            """
            SELECT id, nome, email, role, ativo, created_at
            FROM users
            WHERE organization_id = ?
            ORDER BY id DESC
            LIMIT ?
            """,
            (int(organization_id), int(limit)),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()
