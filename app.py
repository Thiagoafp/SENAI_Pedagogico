"""
app.py
Interface Streamlit para o Sistema Pedagógico SENAI.
Suporta templates personalizados e múltiplos cursos/ementas.
"""

import streamlit as st
import sqlite3
import json
import os
import tempfile
from pathlib import Path

st.set_page_config(
    page_title="SENAI - Sistema Pedagógico",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #003087 0%, #0066CC 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
    }
    .status-gerado { color: #28a745; font-weight: bold; }
</style>
""", unsafe_allow_html=True)


# ─── Login ────────────────────────────────────────────────────────────────────

def check_login():
    """Verifica autenticacao. Retorna True se logado."""
    if st.session_state.get("autenticado"):
        return True

    st.markdown("""
    <div style="max-width:400px;margin:80px auto;padding:32px;
        border-radius:12px;border:1px solid #ddd;background:white;
        box-shadow:0 4px 16px rgba(0,0,0,0.08)">
        <div style="text-align:center;margin-bottom:24px">
            <h2 style="color:#003087;margin:0">Sistema Pedagógico SENAI</h2>
            <p style="color:#116;margin:8px 0 0">Acesso restrito</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("form_login"):
        st.markdown("### Entrar")
        usuario = st.text_input("Usuario")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("Entrar", type="primary", use_container_width=True)

    if entrar:
        try:
            usuario_correto = st.secrets.get("LOGIN_USUARIO", "")
            senha_correta = st.secrets.get("LOGIN_SENHA", "")
        except Exception:
            usuario_correto = ""
            senha_correta = ""

        # Fallback: le do config.json local
        if not senha_correta:
            cfg = carregar_config()
            usuario_correto = cfg.get("login_usuario", "")
            senha_correta = cfg.get("login_senha", "")

        if usuario == usuario_correto and senha == senha_correta and senha_correta:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Usuario ou senha incorretos.")

    return False

CONFIG_FILE = "config.json"
DB_PATH = "senai_ementa.db"
TEMPLATES_DIR = "templates"

TIPOS_DOCUMENTO = {
    "fo_plano":    "📋 FO - Plano de Ensino",
    "apostila":    "📖 Apostila",
    "atividades":  "✏️ Atividades",
    "avaliacao":   "📝 Avaliação",
}

os.makedirs(TEMPLATES_DIR, exist_ok=True)


# ─── Config ───────────────────────────────────────────────────────────────────

def carregar_config() -> dict:
    config = {}
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
    # Streamlit Cloud: usa secrets se api_key nao estiver no config local
    try:
        if not config.get("api_key") and hasattr(st, "secrets"):
            config["api_key"] = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        pass
    return config

def salvar_config(config: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)

def config_valida(config: dict) -> bool:
    return bool(config.get("api_key"))

def drive_configurado(config: dict) -> bool:
    return bool(config.get("gdrive_pasta_raiz_id") and config.get("gdrive_credenciais"))


# ─── Templates ────────────────────────────────────────────────────────────────

def salvar_template(tipo: str, conteudo: bytes):
    """Salva template .docx em disco."""
    path = os.path.join(TEMPLATES_DIR, f"template_{tipo}.docx")
    with open(path, "wb") as f:
        f.write(conteudo)

def carregar_template(tipo: str) -> bytes | None:
    """Carrega template do disco. Retorna None se não existir."""
    path = os.path.join(TEMPLATES_DIR, f"template_{tipo}.docx")
    if os.path.exists(path):
        with open(path, "rb") as f:
            return f.read()
    return None

def remover_template(tipo: str):
    path = os.path.join(TEMPLATES_DIR, f"template_{tipo}.docx")
    if os.path.exists(path):
        os.remove(path)

def tem_template(tipo: str) -> bool:
    return os.path.exists(os.path.join(TEMPLATES_DIR, f"template_{tipo}.docx"))


# ─── Banco ────────────────────────────────────────────────────────────────────

def listar_cursos() -> list:
    if not os.path.exists(DB_PATH):
        return []
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id, nome, ciclo, carga_total FROM cursos ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    return rows

def listar_ucs(curso_id: int) -> list:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, nome, carga_horaria_ha, carga_horaria_hr, serie
        FROM unidades_curriculares WHERE curso_id = ? ORDER BY id
    """, (curso_id,))
    rows = cur.fetchall()
    conn.close()
    return rows

def listar_docs_gerados(uc_id: int) -> dict:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT tipo, drive_id, drive_url, gerado_em
        FROM documentos_gerados WHERE uc_id = ? ORDER BY gerado_em DESC
    """, (uc_id,))
    rows = cur.fetchall()
    conn.close()
    resultado = {}
    for tipo, drive_id, drive_url, gerado_em in rows:
        if tipo not in resultado:
            resultado[tipo] = {"drive_id": drive_id, "drive_url": drive_url, "gerado_em": gerado_em}
    return resultado

def registrar_doc_gerado(uc_id: int, tipo: str, drive_id: str, drive_url: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO documentos_gerados (uc_id, tipo, drive_id, drive_url)
        VALUES (?, ?, ?, ?)
    """, (uc_id, tipo, drive_id, drive_url))
    conn.commit()
    conn.close()


# ─── Geração ──────────────────────────────────────────────────────────────────

def gerar_e_salvar(uc_id: int, uc_nome: str, tipo: str, config: dict,
                   contexto: str = "", nome_docente: str = "") -> tuple:
    """Gera documento. FO usa gerador dedicado. Outros usam template se disponível."""
    
    # FO usa gerador dedicado
    if tipo == "fo_plano":
        from gerador_fo import gerar_fo_completo
        template_fo = "FO_template.docx" if os.path.exists("FO_template.docx") else None
        conteudo_bytes = gerar_fo_completo(
            uc_id, config["api_key"],
            contexto=contexto,
            nome_docente=nome_docente or config.get("nome_professor", ""),
            funcao=config.get("funcao", "Instrutor de Informática"),
            subfuncao=config.get("subfuncao", ""),
            template_path=template_fo
        )
    else:
        from gerador_documentos import gerar_documento
        template_bytes = carregar_template(tipo)
        conteudo_bytes = gerar_documento(uc_id, tipo, config["api_key"], template_bytes)

    # Salva no Drive se configurado
    if drive_configurado(config):
        from gdrive import autenticar_service_account, GDriveManager, nome_arquivo_tipo
        service = autenticar_service_account(config["gdrive_credenciais"])
        drive = GDriveManager(service, config["gdrive_pasta_raiz_id"])
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("""
            SELECT c.nome FROM cursos c
            JOIN unidades_curriculares uc ON uc.curso_id = c.id
            WHERE uc.id = ?
        """, (uc_id,))
        row = cur.fetchone()
        conn.close()
        nome_curso = row[0] if row else "SENAI"
        pasta_uc_id = drive.garantir_estrutura_curso(nome_curso, uc_nome)
        nome_arq = nome_arquivo_tipo(tipo, uc_nome)
        drive_id, drive_url = drive.salvar_docx(conteudo_bytes, nome_arq, pasta_uc_id)
        return drive_id, drive_url, conteudo_bytes

    return None, None, conteudo_bytes


# ─── PÁGINAS ──────────────────────────────────────────────────────────────────

def pagina_configuracoes():
    st.header("⚙️ Configurações")

    config = carregar_config()

    # ── API ──
    with st.form("form_config"):
        st.subheader("🤖 API Claude (Anthropic)")
        api_key = st.text_input("Chave da API", value=config.get("api_key", ""), type="password")

        st.subheader("📁 Google Drive (opcional)")
        gdrive_id = st.text_input("ID da Pasta Raiz", value=config.get("gdrive_pasta_raiz_id", ""))
        gdrive_cred = st.text_input("Caminho credenciais JSON", value=config.get("gdrive_credenciais", "credentials.json"))

        st.subheader("👤 Professor")
        nome_prof = st.text_input("Nome", value=config.get("nome_professor", ""))
        unidade = st.text_input("Unidade SENAI", value=config.get("unidade", ""))
        funcao = st.text_input("Função", value=config.get("funcao", "Instrutor de Informática"),
                               help="Ex: Instrutor de Informática, Professor de TI")
        subfuncao = st.text_input("Subfunção", value=config.get("subfuncao", ""),
                                  help="Ex: Desenvolvimento de Sistemas (opcional)")

        if st.form_submit_button("💾 Salvar", type="primary"):
            salvar_config({
                "api_key": api_key,
                "gdrive_pasta_raiz_id": gdrive_id,
                "gdrive_credenciais": gdrive_cred,
                "nome_professor": nome_prof,
                "unidade": unidade,
                "funcao": funcao,
                "subfuncao": subfuncao,
            })
            st.success("✅ Configurações salvas!")
            st.rerun()

    # ── Templates ──
    st.divider()
    st.subheader("📄 Templates de Documentos")
    st.caption("Faça upload do seu modelo .docx para cada tipo de documento. A IA vai respeitar o seu formato. Deixe em branco para usar o formato padrão profissional.")

    for tipo, label in TIPOS_DOCUMENTO.items():
        col1, col2, col3 = st.columns([3, 2, 1])
        with col1:
            st.markdown(f"**{label}**")
        with col2:
            if tem_template(tipo):
                st.success("✅ Template personalizado ativo")
            else:
                st.info("📝 Usando formato padrão IA")
        with col3:
            if tem_template(tipo):
                if st.button("🗑️ Remover", key=f"rem_{tipo}"):
                    remover_template(tipo)
                    st.rerun()

        arquivo = st.file_uploader(
            f"Upload template {label}",
            type=["docx"],
            key=f"tpl_{tipo}",
            label_visibility="collapsed"
        )
        if arquivo:
            salvar_template(tipo, arquivo.read())
            st.success(f"✅ Template '{label}' salvo!")
            st.rerun()

        st.divider()

    with st.expander("💡 Como usar templates?"):
        st.markdown("""
        1. Abra seu modelo Word atual
        2. Salve como `.docx`
        3. Faça upload aqui
        4. A IA vai ler seu formato e gerar os documentos respeitando sua estrutura
        
        **Não precisa de marcadores especiais** — a IA entende o contexto do seu template automaticamente!
        
        **Dica:** quanto mais completo e detalhado seu template, melhor o resultado!
        """)


def pagina_importar_ementa():
    st.header("📄 Importar Ementa do Curso")
    config = carregar_config()

    if not config.get("api_key"):
        st.warning("⚠️ Configure sua API key antes.")
        return

    col1, col2 = st.columns([2, 1])
    with col1:
        arquivo = st.file_uploader("Selecione o PDF da ementa", type=["pdf"])
    with col2:
        st.info("""
        **O que acontece:**
        1. Extrai texto do PDF
        2. IA identifica todas as UCs
        3. Salva no banco de dados
        
        ⚡ **Só precisa fazer uma vez!**
        """)

    if arquivo:
        st.success(f"✅ {arquivo.name} ({arquivo.size / 1024:.1f} KB)")

        if st.button("🚀 Processar Ementa com IA", type="primary"):
            tmp_path = os.path.join(tempfile.gettempdir(), arquivo.name)
            with open(tmp_path, "wb") as f:
                f.write(arquivo.read())

            with st.status("Processando ementa...", expanded=True) as status:
                try:
                    from extrator_ementa import extrair_texto_pdf, extrair_estrutura_com_ia, salvar_no_banco, criar_banco
                    import anthropic

                    criar_banco()
                    st.write("📖 Extraindo texto do PDF...")
                    texto = extrair_texto_pdf(tmp_path)
                    st.write(f"✅ {len(texto):,} caracteres extraídos")

                    st.write("🤖 Claude analisando a estrutura...")
                    client = anthropic.Anthropic(api_key=config["api_key"])
                    dados = extrair_estrutura_com_ia(texto, client)

                    st.write("💾 Salvando no banco...")
                    salvar_no_banco(dados)

                    ucs = dados.get("unidades_curriculares", [])
                    status.update(label="✅ Ementa processada!", state="complete")
                    st.success(f"**{dados['curso']['nome']}** — {len(ucs)} UCs encontradas!")

                    for i, uc in enumerate(ucs, 1):
                        st.write(f"{i}. **{uc['nome']}** — {uc.get('carga_horaria_ha', '?')} H/A")

                except Exception as e:
                    status.update(label="❌ Erro", state="error")
                    st.error(f"Erro: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())


def pagina_gerar_documentos():
    st.header("📚 Gerar Documentos Pedagógicos")
    config = carregar_config()

    if not config_valida(config):
        st.warning("⚠️ Configure sua API key nas configurações.")
        return

    if not drive_configurado(config):
        st.info("ℹ️ Google Drive não configurado — download local disponível.")

    cursos = listar_cursos()
    if not cursos:
        st.info("📄 Nenhuma ementa importada ainda.")
        return

    # Seleção do curso
    opcoes_cursos = {f"{c[1]} ({c[2] or 'sem ciclo'})": c[0] for c in cursos}
    curso_selecionado = st.selectbox("Selecione o Curso", list(opcoes_cursos.keys()))
    curso_id = opcoes_cursos[curso_selecionado]

    ucs = listar_ucs(curso_id)
    if not ucs:
        st.warning("Nenhuma UC encontrada.")
        return

    st.markdown(f"**{len(ucs)} Unidades Curriculares**")

    # ── Filtro por série ──
    series_disponiveis = sorted(set(u[4] for u in ucs if u[4]))
    serie_filtro = st.selectbox(
        "🎓 Filtrar por série:",
        ["Todas as séries"] + series_disponiveis
    )

    # ── Seleção de UCs ──
    with st.expander("🎯 Selecionar UCs para gerar", expanded=True):
        st.caption("Marque apenas as UCs que precisa agora para economizar tokens.")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Marcar todas"):
                for u in ucs:
                    st.session_state[f"sel_{u[0]}"] = True
                st.rerun()
        with col2:
            if st.button("❌ Desmarcar todas"):
                for u in ucs:
                    st.session_state[f"sel_{u[0]}"] = False
                st.rerun()

        st.divider()
        # Filtra UCs pela série selecionada
        ucs_filtradas = ucs if serie_filtro == "Todas as séries" else [
            u for u in ucs if u[4] == serie_filtro
        ]
        cols = st.columns(2)
        for i, (uc_id_, uc_nome_, ch_ha_, ch_hr_, serie_) in enumerate(ucs_filtradas):
            with cols[i % 2]:
                st.session_state.setdefault(f"sel_{uc_id_}", False)
                st.checkbox(f"{uc_nome_} ({ch_ha_} H/A) — {serie_ or ''}", key=f"sel_{uc_id_}")

    ucs_selecionadas = [u for u in ucs if st.session_state.get(f"sel_{u[0]}", False)]

    if not ucs_selecionadas:
        st.info("☝️ Selecione pelo menos uma UC acima.")
        return

    st.success(f"**{len(ucs_selecionadas)} UC(s) selecionada(s)**")

    # Info sobre templates ativos
    templates_ativos = [t for t in TIPOS_DOCUMENTO if tem_template(t)]
    if templates_ativos:
        labels = [TIPOS_DOCUMENTO[t] for t in templates_ativos]
        st.info(f"🎨 Templates personalizados ativos: {', '.join(labels)}")
    else:
        st.caption("📝 Todos os documentos usarão formato padrão IA")

    st.divider()

    # ── Para cada UC selecionada ──
    for uc_id, uc_nome, ch_ha, ch_hr, serie in ucs_selecionadas:
        docs_gerados = listar_docs_gerados(uc_id)

        with st.expander(f"📚 {uc_nome} — {ch_ha} H/A | {serie or 'Sem série'}"):
            col_info, col_botoes = st.columns([2, 3])

            with col_info:
                st.markdown(f"**Carga:** {ch_ha} H/A / {ch_hr} H/R")
                st.markdown(f"**Série:** {serie or 'N/A'}")
                st.markdown(f"**Gerados:** {len(docs_gerados)}/4")

            with col_botoes:
                cols_btn = st.columns(4)
                for i, (tipo, label) in enumerate(TIPOS_DOCUMENTO.items()):
                    with cols_btn[i]:
                        ja_gerado = tipo in docs_gerados
                        tpl = "🎨" if tem_template(tipo) else "📝"
                        btn_label = f"🔄 {tpl} {label}" if ja_gerado else f"▶️ {tpl} {label}"

                        if ja_gerado and docs_gerados[tipo].get("drive_url"):
                            st.markdown(f"[☁️ Drive]({docs_gerados[tipo]['drive_url']})")

                        if st.button(btn_label, key=f"btn_{tipo}_{uc_id}"):
                            with st.spinner(f"Gerando {label}..."):
                                try:
                                    drive_id, drive_url, conteudo = gerar_e_salvar(
                                        uc_id, uc_nome, tipo, config
                                    )
                                    registrar_doc_gerado(uc_id, tipo, drive_id, drive_url)
                                    st.success(f"✅ {label} gerado!")
                                    st.download_button(
                                        "⬇️ Baixar",
                                        data=conteudo,
                                        file_name=f"{uc_nome} - {serie or ''}.docx".replace("/", "-"),
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"dl_{tipo}_{uc_id}"
                                    )
                                except Exception as e:
                                    st.error(f"❌ {str(e)}")

        # Gerar tudo de uma vez
        if st.button(f"⚡ Gerar TUDO para {uc_nome}", key=f"all_{uc_id}"):
            progress = st.progress(0)
            for i, (tipo, label) in enumerate(TIPOS_DOCUMENTO.items()):
                with st.spinner(f"Gerando {label}..."):
                    try:
                        drive_id, drive_url, _ = gerar_e_salvar(uc_id, uc_nome, tipo, config)
                        registrar_doc_gerado(uc_id, tipo, drive_id, drive_url)
                        progress.progress((i + 1) / len(TIPOS_DOCUMENTO))
                    except Exception as e:
                        st.error(f"❌ {label}: {str(e)}")
            st.success(f"✅ Todos os documentos de '{uc_nome}' gerados!")
            st.rerun()

        st.markdown("---")


def pagina_status():
    st.header("📊 Status dos Documentos")
    cursos = listar_cursos()
    if not cursos:
        st.info("Nenhuma ementa importada.")
        return

    for curso_id, nome, ciclo, carga in cursos:
        st.subheader(f"📘 {nome}")
        st.caption(f"{ciclo or ''} | {carga or ''}")
        ucs = listar_ucs(curso_id)
        dados = []
        total = 0
        for uc_id, uc_nome, ch_ha, ch_hr, serie in ucs:
            docs = listar_docs_gerados(uc_id)
            total += len(docs)
            dados.append({
                "UC": uc_nome, "CH": ch_ha, "Série": serie or "-",
                "Plano": "✅" if "plano_aulas" in docs else "⏳",
                "Apostila": "✅" if "apostila" in docs else "⏳",
                "Atividades": "✅" if "atividades" in docs else "⏳",
                "Avaliação": "✅" if "avaliacao" in docs else "⏳",
            })
        st.dataframe(dados, use_container_width=True)
        total_possivel = len(ucs) * 4
        pct = (total / total_possivel * 100) if total_possivel > 0 else 0
        st.progress(pct / 100, text=f"{total}/{total_possivel} documentos ({pct:.0f}%)")
        st.divider()


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    if not check_login():
        return

    st.markdown("""
    <div class="main-header">
        <h1>🎓 Sistema Pedagógico SENAI</h1>
        <p>Geração Automática de Materiais Didáticos com IA</p>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        config = carregar_config()
        if config_valida(config):
            st.success(f"✅ {config.get('nome_professor', 'Professor')}")
            st.caption(config.get('unidade', ''))
        else:
            st.warning("⚠️ Configure a API key")

        # Templates ativos
        ativos = [TIPOS_DOCUMENTO[t] for t in TIPOS_DOCUMENTO if tem_template(t)]
        if ativos:
            st.caption(f"🎨 Templates: {len(ativos)}/4")

        st.divider()
        if st.button("Sair", use_container_width=True):
            st.session_state["autenticado"] = False
            st.rerun()
        st.divider()
        pagina = st.radio("Navegação", [
            "🏠 Início", "📄 Importar Ementa",
            "📚 Gerar Documentos", "📊 Status", "⚙️ Configurações"
        ], label_visibility="collapsed")

    if pagina == "🏠 Início":
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            ### Como usar:
            1. ⚙️ Configure API key
            2. 📄 Importe a ementa PDF
            3. 📄 (Opcional) Suba seus templates
            4. 📚 Selecione as UCs e gere!
            """)
        with col2:
            st.markdown("""
            ### Documentos:
            - 📋 Plano de Aulas
            - 📖 Apostila
            - ✏️ Atividades
            - 📝 Avaliação
            
            🎨 Com seu template ou formato padrão IA
            """)
        cursos = listar_cursos()
        if cursos:
            st.divider()
            total_ucs = sum(len(listar_ucs(c[0])) for c in cursos)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Cursos", len(cursos))
            c2.metric("UCs", total_ucs)
            c3.metric("Docs possíveis", total_ucs * 4)
            c4.metric("Templates ativos", len([t for t in TIPOS_DOCUMENTO if tem_template(t)]))

    elif pagina == "📄 Importar Ementa":
        pagina_importar_ementa()
    elif pagina == "📚 Gerar Documentos":
        pagina_gerar_documentos()
    elif pagina == "📊 Status":
        pagina_status()
    elif pagina == "⚙️ Configurações":
        pagina_configuracoes()


if __name__ == "__main__":
    main()