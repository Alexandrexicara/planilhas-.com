# 🔧 CORREÇÃO DO ERRO NO RENDER - password_hash

## ❌ Problema Identificado

**Erro:** `no such column: password_hash`

### Causa
O banco de dados no servidor Render tinha a coluna `senha` mas o código estava tentando usar `password_hash`.

**Inconsistência no código:**
- ✅ SELECT usava: `password_hash`
- ❌ INSERT usava: `senha`
- ❌ UPDATE usava: `senha`

---

## ✅ Correção Aplicada

### Arquivo: `web_access_db.py`

#### 1️⃣ Correção no INSERT de usuário
**Antes:**
```sql
INSERT INTO users (organization_id, nome, email, **senha**, role, ativo, created_at)
```

**Depois:**
```sql
INSERT INTO users (organization_id, nome, email, **password_hash**, role, ativo, created_at)
```

#### 2️⃣ Correção no UPDATE de superadmin
**Antes:**
```sql
UPDATE users SET **senha** = ?, role = 'superadmin', ativo = 1
```

**Depois:**
```sql
UPDATE users SET **password_hash** = ?, role = 'superadmin', ativo = 1
```

#### 3️⃣ Correção no INSERT de superadmin
**Antes:**
```sql
INSERT INTO users (organization_id, nome, email, **senha**, role, ativo, created_at)
```

**Depois:**
```sql
INSERT INTO users (organization_id, nome, email, **password_hash**, role, ativo, created_at)
```

---

## 🚀 Como Aplicar no Render

### Opção 1: Deploy Automático (Recomendado)

1. **Commit as alterações:**
   ```bash
   git add web_access_db.py
   git commit -m "fix: corrigir coluna senha para password_hash"
   git push origin main
   ```

2. **O Render fará deploy automático**

3. **Executar script de migração via console do Render:**
   - Acesse: https://dashboard.render.com
   - Selecione seu serviço
   - Vá em "Shell"
   - Execute:
     ```bash
     python corrigir_banco_render.py
     ```

### Opção 2: Migração Manual via Shell

Se o banco já existe no servidor:

```bash
# Acessar shell do Render
python -c "
import sqlite3
conn = sqlite3.connect('acesso_web.db')
cursor = conn.cursor()

# Verificar se precisa migrar
cursor.execute('PRAGMA table_info(users)')
columns = [col[1] for col in cursor.fetchall()]

if 'senha' in columns and 'password_hash' not in columns:
    # Criar nova tabela
    cursor.execute('''
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
    ''')
    
    # Copiar dados
    cursor.execute('''
        INSERT INTO users_new (id, organization_id, nome, email, password_hash, role, ativo, created_at)
        SELECT id, organization_id, nome, email, senha, role, ativo, created_at FROM users
    ''')
    
    # Substituir tabela
    cursor.execute('DROP TABLE users')
    cursor.execute('ALTER TABLE users_new RENAME TO users')
    
    conn.commit()
    print('✅ Migração concluída!')
else:
    print('✅ Banco já está correto!')

conn.close()
"
```

### Opção 3: Recriar Banco (Se não houver dados importantes)

```bash
# No shell do Render
rm acesso_web.db
python app.py
```

O app criará o banco corretamente na inicialização.

---

## ✅ Verificação

Após a correção, teste o login:

1. Acesse: `https://planilhas-1.onrender.com`
2. Faça login com:
   - Email: `admin@planilhas.com`
   - Senha: `admin123`
3. Deve logar sem erros!

---

## 📊 Estrutura Correta da Tabela

```sql
CREATE TABLE users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    organization_id INTEGER,
    nome TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,     -- ✅ Correto!
    role TEXT NOT NULL DEFAULT 'collab',
    ativo INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
);
```

---

## 🎯 Resumo

| Item | Status |
|------|--------|
| Código corrigido | ✅ Sim |
| INSERT corrigido | ✅ Sim |
| UPDATE corrigido | ✅ Sim |
| SELECT (já estava) | ✅ Sim |
| Script de migração | ✅ Criado |
| Documentação | ✅ Completa |

---

## ⚠️ Importante

1. **Faça backup** do banco antes de migrar
2. **Teste localmente** antes de fazer deploy
3. **Verifique os logs** do Render após o deploy
4. **Teste o login** após a migração

---

**Data da Correção**: 24/03/2026  
**Status**: ✅ Código corrigido, pronto para deploy
