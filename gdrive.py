"""
gdrive.py
Integração com Google Drive API.
Salva e organiza os documentos gerados por curso/UC.
"""

import io
import os
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle


SCOPES = ["https://www.googleapis.com/auth/drive"]


def autenticar_service_account(credenciais_json: str):
    """
    Autentica via Service Account (recomendado para produção/Streamlit Cloud).
    credenciais_json: caminho para o arquivo JSON da service account
    """
    creds = service_account.Credentials.from_service_account_file(
        credenciais_json,
        scopes=SCOPES
    )
    service = build("drive", "v3", credentials=creds)
    return service


def autenticar_oauth(credenciais_json: str = "credentials.json", token_pickle: str = "token.pickle"):
    """
    Autentica via OAuth2 (para uso local/desenvolvimento).
    Na primeira vez abre o browser para autorizar.
    """
    creds = None

    if os.path.exists(token_pickle):
        with open(token_pickle, "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credenciais_json, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_pickle, "wb") as token:
            pickle.dump(creds, token)

    service = build("drive", "v3", credentials=creds)
    return service


class GDriveManager:
    """Gerencia pastas e arquivos no Google Drive para o projeto SENAI."""

    def __init__(self, service, pasta_raiz_id: str):
        """
        service: objeto autenticado do Google Drive API
        pasta_raiz_id: ID da pasta raiz configurada pelo professor
        """
        self.service = service
        self.pasta_raiz_id = pasta_raiz_id
        self._cache_pastas = {}  # cache de nome -> id para evitar chamadas repetidas

    def _buscar_ou_criar_pasta(self, nome: str, pai_id: str) -> str:
        """Busca pasta pelo nome dentro de um pai. Cria se não existir."""
        cache_key = f"{pai_id}/{nome}"
        if cache_key in self._cache_pastas:
            return self._cache_pastas[cache_key]

        # Busca existente
        query = (
            f"name = '{nome}' "
            f"and '{pai_id}' in parents "
            f"and mimeType = 'application/vnd.google-apps.folder' "
            f"and trashed = false"
        )
        resultado = self.service.files().list(
            q=query,
            fields="files(id, name)"
        ).execute()

        arquivos = resultado.get("files", [])

        if arquivos:
            pasta_id = arquivos[0]["id"]
        else:
            # Cria a pasta
            metadata = {
                "name": nome,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [pai_id]
            }
            pasta = self.service.files().create(
                body=metadata,
                fields="id"
            ).execute()
            pasta_id = pasta["id"]
            print(f"  📁 Pasta criada: {nome}")

        self._cache_pastas[cache_key] = pasta_id
        return pasta_id

    def garantir_estrutura_curso(self, nome_curso: str, nome_uc: str) -> str:
        """
        Garante que existe a estrutura de pastas:
        pasta_raiz / nome_curso / nome_uc
        Retorna o ID da pasta da UC.
        """
        # Limpa caracteres problemáticos para nome de pasta
        nome_curso_limpo = self._limpar_nome(nome_curso)
        nome_uc_limpo = self._limpar_nome(nome_uc)

        pasta_curso = self._buscar_ou_criar_pasta(nome_curso_limpo, self.pasta_raiz_id)
        pasta_uc = self._buscar_ou_criar_pasta(nome_uc_limpo, pasta_curso)

        return pasta_uc

    def salvar_docx(
        self,
        conteudo_bytes: bytes,
        nome_arquivo: str,
        pasta_id: str
    ) -> tuple[str, str]:
        """
        Salva um arquivo .docx no Drive.
        Retorna (file_id, web_view_link)
        """
        # Verifica se já existe arquivo com esse nome na pasta (atualiza ao invés de duplicar)
        existente_id = self._buscar_arquivo(nome_arquivo, pasta_id)

        media = MediaIoBaseUpload(
            io.BytesIO(conteudo_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            resumable=True
        )

        if existente_id:
            # Atualiza o arquivo existente
            arquivo = self.service.files().update(
                fileId=existente_id,
                media_body=media,
                fields="id, webViewLink"
            ).execute()
            print(f"  🔄 Arquivo atualizado: {nome_arquivo}")
        else:
            # Cria novo arquivo
            metadata = {
                "name": nome_arquivo,
                "parents": [pasta_id]
            }
            arquivo = self.service.files().create(
                body=metadata,
                media_body=media,
                fields="id, webViewLink"
            ).execute()
            print(f"  ✅ Arquivo criado: {nome_arquivo}")

        return arquivo["id"], arquivo.get("webViewLink", "")

    def _buscar_arquivo(self, nome: str, pasta_id: str) -> str | None:
        """Busca arquivo por nome em uma pasta. Retorna ID ou None."""
        query = (
            f"name = '{nome}' "
            f"and '{pasta_id}' in parents "
            f"and trashed = false"
        )
        resultado = self.service.files().list(
            q=query,
            fields="files(id)"
        ).execute()
        arquivos = resultado.get("files", [])
        return arquivos[0]["id"] if arquivos else None

    def listar_arquivos_uc(self, pasta_uc_id: str) -> list:
        """Lista os arquivos já gerados para uma UC."""
        resultado = self.service.files().list(
            q=f"'{pasta_uc_id}' in parents and trashed = false",
            fields="files(id, name, webViewLink, modifiedTime)",
            orderBy="name"
        ).execute()
        return resultado.get("files", [])

    @staticmethod
    def _limpar_nome(nome: str) -> str:
        """Remove caracteres problemáticos para nome de pasta/arquivo."""
        chars_invalidos = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        for char in chars_invalidos:
            nome = nome.replace(char, '-')
        return nome.strip()


def nome_arquivo_tipo(tipo: str, nome_uc: str) -> str:
    """Gera o nome padronizado do arquivo por tipo."""
    nomes = {
        "plano_aulas": "01_Plano_de_Aulas",
        "apostila": "02_Apostila",
        "atividades": "03_Atividades",
        "avaliacao": "04_Avaliacao"
    }
    prefixo = nomes.get(tipo, tipo)
    uc_limpa = GDriveManager._limpar_nome(nome_uc)[:40]  # limita tamanho
    return f"{prefixo}_{uc_limpa}.docx"
