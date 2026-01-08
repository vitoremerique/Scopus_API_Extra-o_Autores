# Projeto Scopus — Extração de Autores por DOI

Este repositório contém um script (`main.py`) que busca metadados de artigos no Scopus (via DOI) e gera arquivos CSV com informações dos autores.

**Requisitos (pip)**
- Python 3.8+ recomendado
- Bibliotecas (instale com pip):

```bash
pip install pandas pybliometrics openpyxl
```

> Observação: `openpyxl` é necessário para o `pandas.read_excel()` ao ler arquivos `.xlsx`.

**Configuração da API Scopus**
1. Crie uma API Key em https://dev.elsevier.com/.
2. Na primeira execução do `main.py`, o script chamará `pybliometrics.scopus.init()` e pedirá a chave.
   - Você também pode configurar manualmente o arquivo de configuração do `pybliometrics` seguindo a documentação da biblioteca.
3. Para extrair os dados completos do Scopus é necessário conectar-se à VPN da USP (acesso institucional).

**Arquivo Excel de entrada**
- Nome padrão esperado: `Scopus Teste.xlsx`
- Colunas esperadas (case-insensitive):
  - `doi` (obrigatório) — coluna contendo os DOIs (pode conter URLs, o script limpa automaticamente)
  - `id` (opcional) — identificador por linha/registro que será usado na saída
- Se preferir outro nome, altere a variável `ARQUIVO_ENTRADA` dentro de `main.py`.

Exemplo: renomeie seu arquivo Excel para exatamente `Scopus Teste.xlsx` e coloque-o na mesma pasta do `main.py`.

**Como executar**
No terminal, na pasta do projeto, execute:

```bash
python main.py
```

- O script lê os DOIs do arquivo Excel, consulta a API do Scopus e gera:
  - `autores_scopus_completo.csv` — registros detalhados por autor
  - `autores_scopus_resumido.csv` — uma linha por artigo com autores formatados



