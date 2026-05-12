"""
Script para corrigir o banco de dados no servidor Render
Executar este script após fazer deploy
"""
import sqlite3
import os

def corrigir_banco_render():
    """Corrige o banco de dados para usar password_hash"""
    print("=" * 60)
    print("🔧 CORRIGINDO BANCO DE DADOS - RENDER")
    print("=" * 60)
    
    db_path = os.path.join(os.path.dirname(__file__), 'acesso_web.db')
    
    if not os.path.exists(db_path):
        print("❌ Banco acesso_web.db não encontrado!")
        return False
    
    print(f"📁 Banco encontrado: {db_path}")
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Verificar estrutura atual
    cursor.execute("PRAGMA table_info(users)")
    columns = cursor.fetchall()
    
    print("\n📊 Colunas atuais da tabela 'users':")
    column_names = [col[1] for col in columns]
    for col in columns:
        print(f"  • {col[1]} ({col[2]})")
    
    # Verificar se precisa de migração
    if 'senha' in column_names and 'password_hash' not in column_names:
        print("\n🔄 Migração necessária: 'senha' → 'password_hash'")
        
        try:
            # Criar nova tabela
            print("1️⃣  Criando nova tabela...")
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
            
            # Copiar dados
            print("2️⃣  Copiando dados...")
            cursor.execute("""
                INSERT INTO users_new (id, organization_id, nome, email, password_hash, role, ativo, created_at)
                SELECT id, organization_id, nome, email, senha, role, ativo, created_at FROM users
            """)
            
            # Remover tabela antiga
            print("3️⃣  Removendo tabela antiga...")
            cursor.execute("DROP TABLE users")
            
            # Renomear nova tabela
            print("4️⃣  Renomeando tabela nova...")
            cursor.execute("ALTER TABLE users_new RENAME TO users")
            
            conn.commit()
            print("\n✅ MIGRAÇÃO CONCLUÍDA COM SUCESSO!")
            
        except Exception as e:
            print(f"\n❌ Erro na migração: {e}")
            conn.rollback()
            return False
    
    elif 'password_hash' in column_names:
        print("\n✅ Banco já está corrigido!")
        print("   Coluna 'password_hash' já existe")
    
    else:
        print("\n⚠️  Estrutura inesperada")
        return False
    
    # Verificar se há usuários
    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]
    print(f"\n👥 Total de usuários: {total_users}")
    
    if total_users > 0:
        cursor.execute("SELECT id, nome, email, role FROM users LIMIT 5")
        users = cursor.fetchall()
        print("\n📋 Usuários encontrados:")
        for user in users:
            print(f"  • ID: {user[0]} | Nome: {user[1]} | Email: {user[2]} | Role: {user[3]}")
    
    conn.close()
    
    print("\n" + "=" * 60)
    print("✅ BANCO CORRIGIDO!")
    print("🚀 Pode fazer login agora!")
    print("=" * 60)
    
    return True

if __name__ == "__main__":
    corrigir_banco_render()
