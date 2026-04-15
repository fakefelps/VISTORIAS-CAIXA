# PreencherDocumentos — Morais Engenharia

Geração automática de **Declaração ART** (Word) e **Memorial** (Excel) em PDF,
a partir dos dados extraídos da ART do CREA-GO.

---

## Estrutura do repositório

```
PreencherDocumentos/
├── app.py                  ← Script principal (UI + lógica)
├── icone.ico               ← Ícone do .exe (adicionar manualmente)
├── README.md
└── .github/
    └── workflows/
        └── build.yml       ← Pipeline CI/CD
```

---

## Como usar o app

1. Baixe o `PreencherDocumentos.exe` na aba **Actions → último build → Artifacts**
2. Execute o `.exe` — não precisa instalar Python nem nenhuma dependência
3. Preencha os campos:
   - Selecione o **template Word** (Declaração ART)
   - Selecione o **Memorial Excel** (versão atual)
   - Selecione a **pasta de assinaturas** (com os arquivos de imagem dos engenheiros)
   - Selecione a **pasta de saída** (onde serão gerados os PDFs)
   - Escolha o **engenheiro** responsável
   - Preencha os **dados da ART** (campos {1} a {11})
   - Marque as **opções**: sistema de esgoto e tipo de lote
   - Informe a **quantidade de casas**
4. Clique em **GERAR DOCUMENTOS**

Os arquivos gerados serão nomeados automaticamente:
- `DECLARAÇÃO ART CS {N} – {endereço}.pdf`
- `MEMORIAL CS {N} – {endereço}.pdf`

---

## Requisitos da máquina do usuário

- Windows 10 ou 11
- Microsoft Word instalado (para exportação PDF)
- Microsoft Excel instalado (para exportação PDF)

---

## Stack técnica

| Componente | Versão |
|---|---|
| Python | 3.11 (build apenas) |
| python-docx | preenchimento Word |
| openpyxl | preenchimento Excel |
| comtypes | exportação PDF via Office COM |
| lxml | manipulação XML checkboxes |
| Pillow | inserção de imagens |
| PyInstaller | empacotamento .exe |

---

## Atualizar o app

1. Edite o arquivo `app.py` no GitHub
2. Faça commit na branch `main`
3. O GitHub Actions gera um novo `.exe` automaticamente (~3 min)
4. Baixe o artefato na aba Actions

---

## Adicionando novo engenheiro

Em `app.py`, localize o dicionário `ENGENHEIROS` e adicione:

```python
"NOME COMPLETO EM MAIÚSCULAS": {
    "cpf": "000.000.000-00",
    "crea": "0000000D-GO",
    "assinatura": "NOME.png",   # nome do arquivo de imagem na pasta assinaturas
},
```
