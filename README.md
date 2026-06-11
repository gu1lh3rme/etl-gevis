# ETL Gevis → VISTORIAS

Automação da atualização da planilha **VISTORIAS.xlsx** a partir de arquivos exportados do sistema web **Gevis**.

## Sumário

- [Visão Geral](#visão-geral)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Execução](#execução)
- [Logs](#logs)
- [Tratamento de Erros](#tratamento-de-erros)
- [Evoluções Futuras](#evoluções-futuras)

---

## Visão Geral

O processo que antes era manual:

1. Exportar tabela do Gevis para Excel
2. Abrir a planilha `VISTORIAS.xlsx`
3. Copiar dados manualmente

Agora é automatizado em um único comando:

```bash
python src/main.py
```

### O que a aplicação faz

| Etapa | Ação |
|-------|------|
| 1 | Lê `data/input/gevis.xlsx` com **pandas** |
| 2 | Lê `templates/VISTORIAS.xlsx` com **openpyxl** |
| 3 | Mapeia colunas conforme `COLUMN_MAPPING` em `config.py` |
| 4 | Preenche colunas fixas (tipo, forma de pagamento, etc.) |
| 5 | Evita duplicatas usando `PLACA/CHASSI` como chave única |
| 6 | Adiciona apenas registros novos ao final da planilha |
| 7 | Salva em `data/output/VISTORIAS_ATUALIZADA.xlsx` sem alterar o original |

---

## Estrutura do Projeto

```
etl-gevis/
│
├── data/
│   ├── input/
│   │   └── gevis.xlsx          ← arquivo exportado do Gevis (não versionado)
│   └── output/
│       └── VISTORIAS_ATUALIZADA.xlsx  ← resultado gerado (não versionado)
│
├── templates/
│   └── VISTORIAS.xlsx          ← planilha modelo (nunca alterada)
│
├── src/
│   ├── main.py         ← ponto de entrada; orquestra o fluxo ETL
│   ├── config.py       ← mapeamentos, caminhos e valores fixos
│   ├── excel_reader.py ← leitura do gevis.xlsx com pandas
│   ├── excel_writer.py ← escrita na planilha com openpyxl
│   ├── mapper.py       ← mapeamento de colunas e valores fixos
│   └── logger.py       ← configuração de logging
│
├── .env                ← variáveis de ambiente (não versionado)
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Requisitos

- Python 3.12+
- pip

---

## Instalação

### 1. Clonar o repositório

```bash
git clone https://github.com/gu1lh3rme/etl-gevis.git
cd etl-gevis
```

### 2. Criar e ativar o ambiente virtual

**Linux / macOS:**
```bash
python -m venv venv
source venv/bin/activate
```

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Instalar dependências

```bash
pip install -r requirements.txt
```

### 4. Adicionar os arquivos de dados

- Exporte a tabela do Gevis e salve como `data/input/gevis.xlsx`
- Coloque a planilha modelo em `templates/VISTORIAS.xlsx`

---

## Configuração

### Mapeamento de colunas

Abra `src/config.py` e ajuste o dicionário `COLUMN_MAPPING`:

```python
COLUMN_MAPPING: dict[str, str] = {
    "Data": "DATA",               # coluna no Gevis → coluna em VISTORIAS
    "Placa": "PLACA/CHASSI",
    "Marca": "VEICULO",
    "Nome Contratante OS": "CLIENTE",
    "Valor": "VLR RECEBIDO",
}
```

> **Chave** = nome exato da coluna no arquivo exportado pelo Gevis  
> **Valor** = nome exato da coluna na planilha VISTORIAS

Se o layout do Gevis mudar (nomes de colunas diferentes), basta atualizar este dicionário.

### Valores fixos

Também em `src/config.py`, o dicionário `FIXED_VALUES` define os valores inseridos automaticamente:

```python
FIXED_VALUES: dict[str, str] = {
    "TIPO DE VISTORIA": "NOVA",
    "FORMA DE PGTO": "PIX",
    "RESP. PELO PGTO": "CLIENTE",
    "STATUS DO PGTO": "PAGO",
}
```

### Caminhos personalizados (opcional)

Copie `.env` e ajuste as variáveis se quiser usar caminhos diferentes dos padrões:

```env
GEVIS_INPUT_PATH=data/input/gevis.xlsx
VISTORIAS_TEMPLATE_PATH=templates/VISTORIAS.xlsx
OUTPUT_PATH=data/output/VISTORIAS_ATUALIZADA.xlsx
```

---

## Execução

```bash
python src/main.py
```

O arquivo de saída será gerado em `data/output/VISTORIAS_ATUALIZADA.xlsx`.

---

## Logs

A aplicação exibe mensagens no terminal:

```
2024-01-10 09:00:01 [INFO] Arquivo Gevis carregado
2024-01-10 09:00:01 [INFO] 15 registros encontrados
2024-01-10 09:00:01 [INFO] 3 registros ignorados por duplicidade
2024-01-10 09:00:01 [INFO] 12 registros adicionados
2024-01-10 09:00:01 [INFO] Arquivo salvo com sucesso: data/output/VISTORIAS_ATUALIZADA.xlsx
```

---

## Tratamento de Erros

| Situação | Mensagem exibida |
|----------|-----------------|
| `gevis.xlsx` não encontrado | `[ERROR] Arquivo Gevis não encontrado: ...` |
| `VISTORIAS.xlsx` não encontrado | `[ERROR] Planilha modelo não encontrada: ...` |
| Coluna ausente no Gevis | `[ERROR] Coluna(s) não encontrada(s) no arquivo Gevis: ['Placa']` |
| Arquivo Gevis vazio | `[ERROR] O arquivo '...' está vazio.` |

Em todos os casos a aplicação encerra com código de saída `1`.

---

## Evoluções Futuras

- **Interface gráfica** com Tkinter para arrastar e soltar arquivos
- **Integração direta** com a API do Gevis, eliminando a exportação manual
- **Agendamento automático** via APScheduler ou cron
- **Exportação para Power BI** via conector REST
- **Relatório de auditoria** em HTML/PDF com os registros inseridos por execução