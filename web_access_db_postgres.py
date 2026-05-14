"""
Banco de dados PostgreSQL para acesso web - Versão Render
Substitui o SQLite volátil por PostgreSQL persistente
"""
import os
# psycopg2 é importado de forma LAZY (só quando DATABASE_URL existir).
# Localmente, sem DATABASE_URL, caímos no fallback SQLite e não precisamos do driver.
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime

def get_db_url():
    """Obtém URL do PostgreSQL do Render"""
    return os.environ.get('DATABASE_URL')

def connect():
    """Conecta ao PostgreSQL ou fallback para SQLite"""
    db_url = get_db_url()
    if not db_url:
        # Fallback para SQLite se DATABASE_URL não existir
        print("DATABASE_URL não encontrado, usando SQLite fallback")
        import sqlite3
        from planilhas_paths import ensure_from_resource
        db_path = ensure_from_resource("acesso_web.db")
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        return conn

    # Import lazy do psycopg2 - só quando realmente vamos usar Postgres
    import psycopg2  # noqa: F401
    conn = psycopg2.connect(db_url)
    conn.autocommit = True
    return conn

def init_db():
    """Inicializa o banco PostgreSQL ou SQLite"""
    conn = connect()
    try:
        db_url = get_db_url()
        is_postgres = bool(db_url)
        
        if is_postgres:
            # PostgreSQL
            with conn.cursor() as cur:
                # Tabela users
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS users (
                        id SERIAL PRIMARY KEY,
                        email VARCHAR(255) UNIQUE NOT NULL,
                        password_hash VARCHAR(255) NOT NULL,
                        nome VARCHAR(255),
                        role VARCHAR(50) DEFAULT 'user',
                        ativo INTEGER DEFAULT 1,
                        organization_id INTEGER,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)

                # Migração automática: se coluna 'senha' existir, renomear
                cur.execute("""
                    SELECT column_name FROM information_schema.columns
                    WHERE table_name='users'
                """)
                cols = [r[0] for r in cur.fetchall()]
                if 'senha' in cols and 'password_hash' not in cols:
                    print("[MIGRAÇÃO] Renomeando 'senha' -> 'password_hash'")
                    cur.execute("ALTER TABLE users RENAME COLUMN senha TO password_hash")
                
                # Tabela organizations
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS organizations (
                        id SERIAL PRIMARY KEY,
                        nome VARCHAR(255) NOT NULL,
                        email VARCHAR(255) NOT NULL,
                        payment_status VARCHAR(50) DEFAULT 'pending',
                        payment_amount DECIMAL(10,2),
                        payment_txid VARCHAR(255),
                        payment_qr_code TEXT,
                        payment_pix_key VARCHAR(255),
                        payment_updated_at TIMESTAMP,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)
                
                # Tabela invites
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS invites (
                        id SERIAL PRIMARY KEY,
                        organization_id INTEGER NOT NULL,
                        email VARCHAR(255),
                        role VARCHAR(50) DEFAULT 'collab',
                        code VARCHAR(255) UNIQUE NOT NULL,
                        used INTEGER DEFAULT 0,
                        used_at TIMESTAMP,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        FOREIGN KEY (organization_id) REFERENCES organizations(id)
                    )
                """)

                # Tabela password_reset_tokens
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS password_reset_tokens (
                        id SERIAL PRIMARY KEY,
                        user_id INTEGER NOT NULL,
                        token VARCHAR(255) UNIQUE NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        expires_at TIMESTAMP NOT NULL,
                        used INTEGER DEFAULT 0,
                        FOREIGN KEY (user_id) REFERENCES users(id)
                    )
                """)
        else:
            # SQLite fallback
            conn.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    nome TEXT,
                    role TEXT DEFAULT 'user',
                    ativo INTEGER DEFAULT 1,
                    organization_id INTEGER,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            conn.execute("""
                CREATE TABLE IF NOT EXISTS organizations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    email TEXT NOT NULL,
                    payment_status TEXT DEFAULT 'pending',
                    payment_amount REAL,
                    payment_txid TEXT,
                    payment_qr_code TEXT,
                    payment_pix_key TEXT,
                    payment_updated_at TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            conn.execute("""
                CREATE TABLE IF NOT EXISTS invites (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    organization_id INTEGER NOT NULL,
                    email TEXT,
                    role TEXT DEFAULT 'collab',
                    code TEXT UNIQUE NOT NULL,
                    used INTEGER DEFAULT 0,
                    used_at TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (organization_id) REFERENCES organizations(id)
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS password_reset_tokens (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    token TEXT UNIQUE NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    expires_at TEXT NOT NULL,
                    used INTEGER DEFAULT 0,
                    FOREIGN KEY (user_id) REFERENCES users(id)
                )
            """)
            
            conn.commit()
        
        print(f"Banco {'PostgreSQL' if is_postgres else 'SQLite'} inicializado com sucesso!")
    finally:
        conn.close()

def create_user(organization_id, nome, email, senha, role="collab", ativo=1):
    """Cria usuário no PostgreSQL ou SQLite"""
    conn = connect()
    try:
        password_hash = generate_password_hash(senha)
        db_url = get_db_url()
        is_postgres = bool(db_url)
        
        if is_postgres:
            # PostgreSQL
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO users (organization_id, nome, email, password_hash, role, ativo, created_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    RETURNING id
                """, (organization_id, nome.strip(), email.strip().lower(), password_hash, role, int(ativo), datetime.utcnow()))
                
                user_id = cur.fetchone()[0]
        else:
            # SQLite
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO users (organization_id, nome, email, password_hash, role, ativo, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (organization_id, nome.strip(), email.strip().lower(), password_hash, role, int(ativo), datetime.utcnow()))
            
            conn.commit()
            user_id = cur.lastrowid
        
        print(f"Usuário criado com ID: {user_id}")
        return user_id
    finally:
        conn.close()

def authenticate(email, senha):
    """Autentica usuário no PostgreSQL ou SQLite"""
    print(f"=== DEBUG AUTHENTICATE ===")
    print(f"Email recebido: '{email}'")
    print(f"Senha recebida: '{senha[:3]}***' (len={len(senha)})")
    
    if not email or not senha:
        print("Email ou senha vazios")
        return None

    conn = connect()
    try:
        db_url = get_db_url()
        is_postgres = bool(db_url)
        print(f"Usando PostgreSQL: {is_postgres}")
        
        # Listar todos os usuários para debug
        if is_postgres:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("SELECT id, email, role, ativo FROM users ORDER BY id")
                all_users = cur.fetchall()
                print(f"POSTGRES - Todos usuários: {all_users}")
        else:
            all_users = conn.execute("SELECT id, email, role, ativo FROM users ORDER BY id").fetchall()
            print(f"SQLITE - Todos usuários: {all_users}")
        
        if is_postgres:
            # PostgreSQL
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute("""
                    SELECT id, organization_id, nome, email, password_hash, role, ativo
                    FROM users
                    WHERE email = %s
                """, (email.strip().lower(),))
                
                row = cur.fetchone()
        else:
            # SQLite
            row = conn.execute("""
                SELECT id, organization_id, nome, email, password_hash, role, ativo
                FROM users
                WHERE email = ?
            """, (email.strip().lower(),)).fetchone()
        
        print(f"Usuário encontrado: {row}")
        
        if not row:
            print("Usuário não encontrado")
            return None
            
        if int(row["ativo"]) != 1:
            print(f"Usuário inativo (ativo={row['ativo']})")
            return None

        print(f"Verificando senha...")
        print(f"Hash armazenado: {row['password_hash'][:50]}...")
        print(f"Senha fornecida: '{senha}'")
        
        if not check_password_hash(row["password_hash"], senha):
            print("❌ Senha incorreta")
            return None

        print("✅ Autenticação bem-sucedida!")
        return {
            "id": row["id"],
            "organization_id": row["organization_id"],
            "nome": row["nome"],
            "email": row["email"],
            "role": row["role"],
        }
    finally:
        conn.close()

def ensure_superadmin(email, senha):
    """Garante que superadmin exista no PostgreSQL"""
    if not email or not senha:
        return None

    email_norm = email.strip().lower()
    conn = connect()
    try:
        with conn.cursor() as cur:
            # Verificar se já existe
            cur.execute("SELECT id FROM users WHERE email = %s", (email_norm,))
            existing = cur.fetchone()
            
            password_hash = generate_password_hash(senha)
            
            if existing:
                # Atualizar existente
                cur.execute("""
                    UPDATE users
                    SET password_hash = %s, role = 'superadmin', ativo = 1
                    WHERE id = %s
                """, (password_hash, existing[0]))
                print(f"Superadmin atualizado com ID: {existing[0]}")
                return existing[0]
            else:
                # Criar novo
                cur.execute("""
                    INSERT INTO users (organization_id, nome, email, password_hash, role, ativo, created_at)
                    VALUES (NULL, 'Criador', %s, %s, 'superadmin', 1, %s)
                    RETURNING id
                """, (email_norm, password_hash, datetime.utcnow()))
                
                user_id = cur.fetchone()[0]
                print(f"Superadmin criado com ID: {user_id}")
                return user_id
    finally:
        conn.close()

def get_user(user_id):
    """Obtém usuário por ID"""
    conn = connect()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, organization_id, nome, email, role, ativo
                FROM users
                WHERE id = %s
            """, (user_id,))
            
            return cur.fetchone()
    finally:
        conn.close()

def create_organization(nome, email, payment_amount=None):
    """Cria organização no PostgreSQL"""
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO organizations (nome, email, payment_amount, created_at)
                VALUES (%s, %s, %s, %s)
                RETURNING id
            """, (nome.strip(), email.strip().lower(), payment_amount, datetime.utcnow()))
            
            org_id = cur.fetchone()[0]
            print(f"Organização criada com ID: {org_id}")
            return org_id
    finally:
        conn.close()

def get_organization(org_id):
    """Obtém organização por ID"""
    conn = connect()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, nome, email, payment_status, payment_amount, payment_txid, payment_qr_code, payment_pix_key, created_at
                FROM organizations
                WHERE id = %s
            """, (org_id,))
            
            return cur.fetchone()
    finally:
        conn.close()

def organization_has_access(org_id):
    """Verifica se organização tem acesso.
    
    [LIBERADO] Pagamento desativado temporariamente: sempre retorna True
    enquanto a organização existir. Para reativar exigência, restaurar a
    consulta a payment_status == 'paid'.
    """
    if not org_id:
        return False
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT id FROM organizations WHERE id = %s", (org_id,))
            return cur.fetchone() is not None
    finally:
        conn.close()

def list_users(org_id):
    """Lista usuários de uma organização"""
    conn = connect()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, nome, email, role, ativo, created_at
                FROM users
                WHERE organization_id = %s
                ORDER BY created_at DESC
            """, (org_id,))
            
            return cur.fetchall()
    finally:
        conn.close()

def set_organization_payment_pending(org_id, txid, qr_code, pix_key):
    """Define pagamento como pendente"""
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE organizations
                SET payment_status = 'pending', payment_txid = %s, payment_qr_code = %s, payment_pix_key = %s, payment_updated_at = %s
                WHERE id = %s
            """, (txid, qr_code, pix_key, datetime.utcnow(), org_id))
    finally:
        conn.close()

def set_organization_paid(org_id):
    """Define organização como paga"""
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE organizations
                SET payment_status = 'paid', payment_updated_at = %s
                WHERE id = %s
            """, (datetime.utcnow(), org_id))
    finally:
        conn.close()

def create_invite(org_id, email, role="collab"):
    """Cria convite"""
    import secrets
    code = secrets.token_urlsafe(8)
    
    conn = connect()
    try:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO invites (organization_id, email, role, code, created_at)
                VALUES (%s, %s, %s, %s, %s)
                RETURNING id
            """, (org_id, email.strip().lower(), role, code, datetime.utcnow()))
            
            return cur.fetchone()[0], code
    finally:
        conn.close()

def redeem_invite(code, nome, email, senha):
    """Resgata convite"""
    conn = connect()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT i.id, i.organization_id, i.email, i.role
                FROM invites i
                WHERE i.code = %s AND i.used = 0
            """, (code,))
            
            invite = cur.fetchone()
            if not invite:
                return False, "Convite inválido ou já usado", None
            
            if invite["email"] and invite["email"] != email.strip().lower():
                return False, "Este convite é para outro email", None
            
            # Criar usuário
            user_id = create_user(invite["organization_id"], nome, email.strip().lower(), senha, invite["role"])
            
            # Marcar convite como usado
            cur.execute("""
                UPDATE invites
                SET used = 1, used_at = %s
                WHERE id = %s
            """, (datetime.utcnow(), invite["id"]))
            
            return True, "Usuário criado com sucesso", user_id
    finally:
        conn.close()

def list_invites(org_id):
    """Lista convites de uma organização"""
    conn = connect()
    try:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("""
                SELECT id, email, role, code, used, used_at, created_at
                FROM invites
                WHERE organization_id = %s
                ORDER BY created_at DESC
            """, (org_id,))
            
            return cur.fetchall()
    finally:
        conn.close()


def create_password_reset_token(email, ttl_minutes=30):
    """Gera token de reset e retorna (user_dict, token) ou None."""
    import secrets
    from datetime import timedelta
    conn = connect()
    try:
        db_url = get_db_url()
        is_postgres = bool(db_url)
        token = secrets.token_urlsafe(32)
        now = datetime.utcnow()
        expires = now + timedelta(minutes=ttl_minutes)

        if is_postgres:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute(
                    "SELECT id, nome, email FROM users WHERE lower(email) = %s AND ativo = 1",
                    (email.strip().lower(),),
                )
                row = cur.fetchone()
                if not row:
                    return None
                cur.execute(
                    """
                    INSERT INTO password_reset_tokens (user_id, token, created_at, expires_at, used)
                    VALUES (%s, %s, %s, %s, 0)
                    """,
                    (row["id"], token, now, expires),
                )
                return ({"id": row["id"], "nome": row["nome"], "email": row["email"]}, token)
        else:
            cur = conn.execute(
                "SELECT id, nome, email FROM users WHERE lower(email) = ? AND ativo = 1",
                (email.strip().lower(),),
            )
            row = cur.fetchone()
            if not row:
                return None
            conn.execute(
                """
                INSERT INTO password_reset_tokens (user_id, token, created_at, expires_at, used)
                VALUES (?, ?, ?, ?, 0)
                """,
                (row["id"], token, now.strftime("%Y-%m-%d %H:%M:%S"), expires.strftime("%Y-%m-%d %H:%M:%S")),
            )
            conn.commit()
            return ({"id": row["id"], "nome": row["nome"], "email": row["email"]}, token)
    finally:
        conn.close()


def consume_reset_token_and_update_password(token, nova_senha):
    """Valida token, troca senha, marca como usado. Retorna True/False."""
    conn = connect()
    try:
        db_url = get_db_url()
        is_postgres = bool(db_url)
        new_hash = generate_password_hash(nova_senha)
        now = datetime.utcnow()

        if is_postgres:
            with conn.cursor(cursor_factory=RealDictCursor) as cur:
                cur.execute(
                    "SELECT user_id, expires_at, used FROM password_reset_tokens WHERE token = %s",
                    (token,),
                )
                row = cur.fetchone()
                if not row or row["used"]:
                    return False
                if row["expires_at"] < now:
                    return False
                cur.execute("UPDATE users SET password_hash = %s WHERE id = %s", (new_hash, row["user_id"]))
                cur.execute("UPDATE password_reset_tokens SET used = 1 WHERE token = %s", (token,))
                return True
        else:
            cur = conn.execute(
                "SELECT user_id, expires_at, used FROM password_reset_tokens WHERE token = ?",
                (token,),
            )
            row = cur.fetchone()
            if not row or row["used"]:
                return False
            try:
                exp = datetime.strptime(row["expires_at"], "%Y-%m-%d %H:%M:%S")
            except Exception:
                return False
            if datetime.utcnow() > exp:
                return False
            conn.execute("UPDATE users SET password_hash = ? WHERE id = ?", (new_hash, row["user_id"]))
            conn.execute("UPDATE password_reset_tokens SET used = 1 WHERE token = ?", (token,))
            conn.commit()
            return True
    finally:
        conn.close()


print("Módulo PostgreSQL carregado com sucesso!")
