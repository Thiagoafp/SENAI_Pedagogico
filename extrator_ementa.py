"""
extrator_ementa.py
Lê a ementa PDF do curso SENAI e salva estruturado no SQLite.
Só precisa rodar uma vez por ementa!
"""

import sqlite3
import json
import re
import pdfplumber
import anthropic
from pathlib import Path


# ─── Configurações ────────────────────────────────────────────────────────────
DB_PATH = "senai_ementa.db"
# ──────────────────────────────────────────────────────────────────────────────


def criar_banco():
    """Cria as tabelas no SQLite se não existirem."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.executescript("""
        CREATE TABLE IF NOT EXISTS cursos (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nome        TEXT NOT NULL,
            ciclo       TEXT,
            carga_total TEXT,
            objetivo    TEXT,
            criado_em   DATETIME DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS unidades_curriculares (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            curso_id        INTEGER REFERENCES cursos(id),
            nome            TEXT NOT NULL,
            carga_horaria_ha TEXT,
            carga_horaria_hr TEXT,
            serie           TEXT,
            objetivo_geral  TEXT
        );

        CREATE TABLE IF NOT EXISTS capacidades (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            uc_id   INTEGER REFERENCES unidades_curriculares(id),
            descricao TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS conhecimentos (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            uc_id   INTEGER REFERENCES unidades_curriculares(id),
            descricao TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS referencias (
            id      INTEGER PRIMARY KEY AUTOINCREMENT,
            uc_id   INTEGER REFERENCES unidades_curriculares(id),
            descricao TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS documentos_gerados (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            uc_id       INTEGER REFERENCES unidades_curriculares(id),
            tipo        TEXT NOT NULL,  -- 'plano_aulas', 'apostila', 'atividades', 'avaliacao'
            drive_id    TEXT,
            drive_url   TEXT,
            gerado_em   DATETIME DEFAULT CURRENT_TIMESTAMP
        );
    """)

    conn.commit()
    conn.close()
    print("✅ Banco criado/verificado com sucesso!")


def extrair_texto_pdf(caminho_pdf: str) -> str:
    """Extrai todo o texto do PDF usando pdfplumber."""
    texto_completo = []
    with pdfplumber.open(caminho_pdf) as pdf:
        print(f"📄 PDF com {len(pdf.pages)} páginas")
        for i, page in enumerate(pdf.pages):
            texto = page.extract_text()
            if texto:
                texto_completo.append(f"[PÁGINA {i+1}]\n{texto}")
    return "\n\n".join(texto_completo)


def extrair_estrutura_com_ia(texto_pdf: str, client: anthropic.Anthropic) -> dict:
    """
    Usa Claude para extrair a estrutura da ementa em JSON.
    Chamada única - economiza tokens nas gerações futuras!
    """
    print("🤖 Enviando para Claude extrair estrutura...")

    prompt = f"""Você é um especialista em análise de ementas de cursos técnicos SENAI.

Analise o texto abaixo de uma ementa de curso técnico e extraia as informações estruturadas.

RETORNE APENAS JSON VÁLIDO, sem markdown, sem explicações, sem ```json.

REGRA IMPORTANTE: Se uma UC pertence a múltiplas séries (ex: "1ª, 2ª e 3ª série"), 
crie UMA ENTRADA POR SÉRIE com a carga horária dividida igualmente.
Nunca coloque múltiplas séries em uma única UC.

Estrutura esperada:
{{
  "curso": {{
    "nome": "nome do curso",
    "ciclo": "2º ciclo ou similar",
    "carga_total": "carga horária total",
    "objetivo": "objetivo geral do curso"
  }},
  "unidades_curriculares": [
    {{
      "nome": "nome da UC",
      "carga_horaria_ha": "120",
      "carga_horaria_hr": "100", 
      "serie": "1ª série",
      "IMPORTANTE_SERIES": "Se a UC aparece em MAIS DE UMA série (ex: 1ª, 2ª e 3ª série), crie uma entrada SEPARADA para cada série com carga_horaria_ha dividida igualmente. Ex: UC com 120 H/A em 3 séries = 3 entradas de 40 H/A cada uma.",
      "objetivo_geral": "competência/objetivo da UC",
      "capacidades": [
        "capacidade 1",
        "capacidade 2"
      ],
      "conhecimentos": [
        "conhecimento 1",
        "conhecimento 2"
      ],
      "referencias": [
        "referência bibliográfica 1"
      ]
    }}
  ]
}}

TEXTO DA EMENTA:
{texto_pdf[:80000]}
"""

    message = client.messages.create(
       ## model="claude-sonnet-4-20250514",
        model="claude-sonnet-4-5",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}]
    )

    resposta = message.content[0].text.strip()

    # Remove possíveis marcadores de markdown se vieram mesmo assim
    resposta = re.sub(r'^```json\s*', '', resposta)
    resposta = re.sub(r'\s*```$', '', resposta)

    return json.loads(resposta)


def salvar_no_banco(dados: dict) -> int:
    """Salva a estrutura extraída no SQLite. Retorna o ID do curso."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # Salva o curso
    curso = dados["curso"]
    cur.execute("""
        INSERT INTO cursos (nome, ciclo, carga_total, objetivo)
        VALUES (?, ?, ?, ?)
    """, (
        curso.get("nome", ""),
        curso.get("ciclo", ""),
        curso.get("carga_total", ""),
        curso.get("objetivo", "")
    ))
    curso_id = cur.lastrowid
    print(f"✅ Curso salvo: {curso['nome']} (ID: {curso_id})")

    # Salva cada UC
    for uc_data in dados.get("unidades_curriculares", []):
        cur.execute("""
            INSERT INTO unidades_curriculares 
            (curso_id, nome, carga_horaria_ha, carga_horaria_hr, serie, objetivo_geral)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            curso_id,
            uc_data.get("nome", ""),
            uc_data.get("carga_horaria_ha", ""),
            uc_data.get("carga_horaria_hr", ""),
            uc_data.get("serie", ""),
            uc_data.get("objetivo_geral", "")
        ))
        uc_id = cur.lastrowid

        # Capacidades
        for cap in uc_data.get("capacidades", []):
            if cap.strip():
                cur.execute(
                    "INSERT INTO capacidades (uc_id, descricao) VALUES (?, ?)",
                    (uc_id, cap.strip())
                )

        # Conhecimentos
        for con in uc_data.get("conhecimentos", []):
            if con.strip():
                cur.execute(
                    "INSERT INTO conhecimentos (uc_id, descricao) VALUES (?, ?)",
                    (uc_id, con.strip())
                )

        # Referências
        for ref in uc_data.get("referencias", []):
            if ref.strip():
                cur.execute(
                    "INSERT INTO referencias (uc_id, descricao) VALUES (?, ?)",
                    (uc_id, ref.strip())
                )

        print(f"  📚 UC salva: {uc_data['nome']}")

    conn.commit()
    conn.close()
    return curso_id


def carregar_uc(uc_id: int) -> dict:
    """
    Carrega todos os dados de uma UC do banco.
    Usado na geração de documentos - token mínimo!
    """
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("SELECT * FROM unidades_curriculares WHERE id = ?", (uc_id,))
    row = cur.fetchone()
    if not row:
        return {}

    colunas = [d[0] for d in cur.description]
    uc = dict(zip(colunas, row))

    cur.execute("SELECT descricao FROM capacidades WHERE uc_id = ?", (uc_id,))
    uc["capacidades"] = [r[0] for r in cur.fetchall()]

    cur.execute("SELECT descricao FROM conhecimentos WHERE uc_id = ?", (uc_id,))
    uc["conhecimentos"] = [r[0] for r in cur.fetchall()]

    cur.execute("SELECT descricao FROM referencias WHERE uc_id = ?", (uc_id,))
    uc["referencias"] = [r[0] for r in cur.fetchall()]

    conn.close()
    return uc


def listar_ucs(curso_id: int = None) -> list:
    """Lista todas as UCs, opcionalmente filtrando por curso."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    if curso_id:
        cur.execute("""
            SELECT uc.id, uc.nome, uc.carga_horaria_ha, uc.serie, c.nome as curso
            FROM unidades_curriculares uc
            JOIN cursos c ON c.id = uc.curso_id
            WHERE uc.curso_id = ?
            ORDER BY uc.id
        """, (curso_id,))
    else:
        cur.execute("""
            SELECT uc.id, uc.nome, uc.carga_horaria_ha, uc.serie, c.nome as curso
            FROM unidades_curriculares uc
            JOIN cursos c ON c.id = uc.curso_id
            ORDER BY c.id, uc.id
        """)

    rows = cur.fetchall()
    conn.close()
    return rows


def listar_cursos() -> list:
    """Lista todos os cursos no banco."""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id, nome, ciclo, carga_total FROM cursos ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    return rows


def processar_ementa(caminho_pdf: str, api_key: str):
    """
    Fluxo completo: PDF → texto → IA → SQLite
    Chame esta função uma vez por ementa!
    """
    print(f"\n🚀 Processando ementa: {caminho_pdf}")
    print("=" * 50)

    # 1. Criar banco
    criar_banco()

    # 2. Extrair texto do PDF
    print("\n📖 Extraindo texto do PDF...")
    texto = extrair_texto_pdf(caminho_pdf)
    print(f"✅ {len(texto)} caracteres extraídos")

    # 3. IA extrai estrutura
    client = anthropic.Anthropic(api_key=api_key)
    dados = extrair_estrutura_com_ia(texto, client)

    # 4. Salvar no banco
    print("\n💾 Salvando no banco de dados...")
    curso_id = salvar_no_banco(dados)

    # 5. Resumo
    print("\n" + "=" * 50)
    print("✅ PROCESSAMENTO CONCLUÍDO!")
    print(f"   Curso ID: {curso_id}")
    print(f"   UCs encontradas: {len(dados.get('unidades_curriculares', []))}")
    print(f"   Banco salvo em: {DB_PATH}")
    print("\nPróximo passo: use o Streamlit para gerar os documentos!")

    return curso_id


# ─── Execução direta (teste) ───────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    import os

    if len(sys.argv) < 2:
        print("Uso: python extrator_ementa.py caminho_ementa.pdf")
        print("\nOu importe e use as funções:")
        print("  from extrator_ementa import processar_ementa")
        print("  processar_ementa('ementa.pdf', 'sua-api-key')")
        sys.exit(0)

    caminho = sys.argv[1]
    key = os.environ.get("ANTHROPIC_API_KEY", "")

    if not key:
        print("❌ Configure ANTHROPIC_API_KEY no ambiente")
        sys.exit(1)

    processar_ementa(caminho, key)