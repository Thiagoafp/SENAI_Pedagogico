# 🎓 Sistema Pedagógico SENAI

Geração automática de materiais didáticos com IA (Claude) a partir da ementa do curso.

## Arquitetura

```
senai_pedagogico/
├── app.py                  # Interface Streamlit principal
├── extrator_ementa.py      # PDF → SQLite (roda uma vez por ementa)
├── gerador_documentos.py   # Claude API → docx (plano, apostila, atividades, avaliação)
├── gdrive.py               # Integração Google Drive API
├── requirements.txt
└── senai_ementa.db         # SQLite gerado automaticamente
```

## Fluxo de dados

```
PDF da Ementa
      ↓ (uma vez só)
extrator_ementa.py
      ↓
SQLite (cursos, UCs, capacidades, conhecimentos, referências)
      ↓ (por demanda, token mínimo)
gerador_documentos.py + Claude API
      ↓
Arquivo .docx
      ↓
gdrive.py → Google Drive (organizado por curso/UC)
```

## Instalação

```bash
pip install -r requirements.txt
```

## Configuração Google Drive

1. Acesse [Google Cloud Console](https://console.cloud.google.com)
2. Crie um projeto e ative a **Google Drive API**
3. Crie uma **Service Account** → baixe o JSON
4. Compartilhe sua pasta do Drive com o email da Service Account
5. Cole o caminho do JSON e o ID da pasta no app

## Uso

```bash
streamlit run app.py
```

### Passo a passo:
1. ⚙️ Configure API key e Google Drive
2. 📄 Importe a ementa PDF (processa uma vez, salva no SQLite)
3. 📚 Selecione uma UC e gere os documentos
4. 📊 Acompanhe o status de geração

## Documentos gerados por UC

| Arquivo | Conteúdo |
|---------|----------|
| `01_Plano_de_Aulas_[UC].docx` | Sequência didática aula a aula |
| `02_Apostila_[UC].docx` | Material teórico completo |
| `03_Atividades_[UC].docx` | Exercícios e situações-problema |
| `04_Avaliacao_[UC].docx` | Prova formal com gabarito |

## Estrutura de pastas no Drive

```
📁 Pasta Raiz (configurada)
  📁 Técnico em Desenvolvimento de Sistemas
    📁 Lógica de Programação I
      📄 01_Plano_de_Aulas_...docx
      📄 02_Apostila_...docx
      📄 03_Atividades_...docx
      📄 04_Avaliacao_...docx
    📁 Banco de Dados I
      📄 ...
```

## Para outros professores

Cada professor:
1. Baixa o projeto
2. Configura sua própria API key
3. Configura sua própria pasta do Drive (ID)
4. Importa a ementa do seu curso
5. Gera os materiais!
