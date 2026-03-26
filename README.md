# 📋 DUIMP → DI Converter

Sistema completo para converter arquivos da **DUIMP** (Declaração Única de Importação)
no formato estruturado da antiga **DI (Declaração de Importação)** — com layout profissional
em Excel, cálculo automático de impostos e interface web via Flask.

---

## 📁 Estrutura do Projeto

```
duimp_converter/
│
├── app.py              ← Servidor Flask (rotas, upload, download)
├── parser.py           ← Leitura de XML, JSON e Excel (DUIMP)
├── processor.py        ← Lógica de negócio, cálculo de impostos
├── generator.py        ← Geração do Excel no formato DI
│
├── templates/
│   └── index.html      ← Interface web de upload
│
├── uploads/            ← Arquivos temporários de entrada (auto-limpeza)
├── outputs/            ← Arquivos Excel gerados (auto-limpeza)
│
├── exemplo_duimp.xml   ← Arquivo de teste fictício (14 adições)
├── requirements.txt    ← Dependências Python
└── README.md           ← Este arquivo
```

---

## ⚙️ Pré-requisitos

- **Python 3.9+** instalado
- **pip** disponível no terminal

Verifique com:
```bash
python --version
pip --version
```

---

## 🚀 Instalação Passo a Passo

### 1. Clone ou baixe o projeto

```bash
# Se tiver git:
git clone https://github.com/seu-usuario/duimp-di-converter.git
cd duimp-di-converter

# Ou descompacte o ZIP e entre na pasta:
cd duimp_converter
```

### 2. Crie um ambiente virtual (recomendado)

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux / macOS:**
```bash
python -m venv venv
source venv/bin/activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

### 4. Inicie o servidor

```bash
python app.py
```

Você verá:
```
============================================================
  DUIMP to DI Converter – Iniciando servidor...
  Acesse: http://localhost:5000
============================================================
```

### 5. Acesse no navegador

```
http://localhost:5000
```

---

## 📂 Como Usar

### Via Interface Web

1. Abra `http://localhost:5000`
2. Clique em **"Selecionar arquivo"** ou arraste o arquivo DUIMP
3. *(Opcional)* Informe a taxa de câmbio manualmente (ex: `5.2440`)
4. Clique em **"Converter e Baixar Excel"**
5. O arquivo `.xlsx` será baixado automaticamente

### Via API (cURL / Postman)

```bash
# Enviar arquivo e receber JSON com resumo
curl -X POST http://localhost:5000/api/converter \
  -F "file=@exemplo_duimp.xml" \
  -F "taxa_cambio=5.244"
```

Resposta:
```json
{
  "status": "success",
  "numero_duimp": "26BR0000295202-6",
  "total_adicoes": 14,
  "totais": {
    "valor_aduaneiro": 29304.43,
    "ii": 5274.81,
    "ipi": 1728.99,
    "pis": 762.51,
    "cofins": 3503.76,
    "icms": 8906.61,
    "total_tributos": 20176.68
  },
  "download_id": "abc12345"
}
```

Download pelo ID:
```bash
curl -O http://localhost:5000/download/abc12345
```

---

## 📄 Formatos de Entrada Aceitos

### XML (recomendado)
Use a estrutura do arquivo `exemplo_duimp.xml` como modelo.
Campos principais:

| Tag XML              | Descrição                    |
|----------------------|------------------------------|
| `<numeroDuimp>`      | Número da DUIMP              |
| `<cnpjImportador>`   | CNPJ do importador           |
| `<taxaCambio>`       | Taxa de câmbio USD → BRL     |
| `<adicao>`           | Bloco de cada adição         |
| `<ncm>`              | NCM da mercadoria            |
| `<valorCondicaoVenda>`| Valor FOB em USD            |
| `<iiAliquota>`       | Alíquota II (%)              |
| `<ipiAliquota>`      | Alíquota IPI (%)             |

### JSON
Estrutura equivalente ao XML, com chaves em camelCase:
```json
{
  "duimp": {
    "cabecalho": {
      "numeroDuimp": "26BR0000295202-6",
      "taxaCambio": 5.244
    },
    "adicoes": [
      {
        "ncm": "8512.2011",
        "descricao": "FAROL DIREITO FIAT",
        "valorCondicaoVenda": 732.00,
        "iiAliquota": 18.0
      }
    ]
  }
}
```

### Excel (.xlsx)
- Aba **"Cabecalho"** com pares chave/valor (coluna A = rótulo, coluna B = valor)
- Aba **"Adicoes"** com tabela onde as colunas correspondem aos campos:
  `adicao | ncm | descricao | quantidade | valor_total_usd | ii_aliquota | ipi_aliquota | ...`

---

## 📊 Saída Excel – 3 Abas

### Aba 1: "Extrato DI"
Formato semelhante ao extrato da Receita Federal:
- Cabeçalho com dados do processo
- Resumo de valores (FOB, Frete, Valor Aduaneiro)
- Tabela de todas as adições com impostos
- Linha de totais consolidada

### Aba 2: "Adições Detalhadas"
Tabela completa com todos os campos por adição:
- Fabricante, país de origem, detalhamento do produto
- Alíquotas e valores separados para cada tributo
- Peso líquido, quantidade, condição

### Aba 3: "Resumo Tributário"
Dashboard fiscal:
- Total por tributo em R$
- Percentual de cada tributo sobre o valor aduaneiro
- Nota de rodapé com observações

---

## 🧮 Cálculo Automático de Impostos

Quando os impostos **não vierem preenchidos** no arquivo, o sistema calcula automaticamente:

| Tributo | Lógica de Cálculo                        | Alíquota Padrão   |
|---------|------------------------------------------|-------------------|
| **II**  | `Valor Aduaneiro × Alíq%`               | 18% (autopeças)   |
| **IPI** | `(VA + II) × Alíq%`                     | 5% (cap. 8512/8708)|
| **PIS** | `(VA + II + IPI) × 2,10%`              | 2,10%             |
| **COFINS** | `(VA + II + IPI) × 9,65%`           | 9,65%             |
| **ICMS** | `(VA+II+IPI+PIS+COFINS) / (1-18%) × 18%` | 18% (SC)       |

> **Nota:** O Valor Aduaneiro por adição é calculado como `Valor USD × Taxa de Câmbio`
> quando não informado diretamente.

---

## 🔍 Descrição de Cada Arquivo

### `parser.py`
Responsável pela leitura do arquivo de entrada.
- `parse_file(filepath)` — detecta o formato e delega ao parser correto
- `_parse_xml()` — lê estrutura XML com remoção de namespaces
- `_parse_json()` — lê JSON com mapeamento flexível de chaves
- `_parse_excel()` — lê Excel com suporte a múltiplas abas e variações de nome de coluna
- Todas as funções são tolerantes a campos ausentes (retornam `""` ou `0.0`)

### `processor.py`
Aplica a lógica fiscal sobre os dados brutos.
- `process_data(raw, taxa_cambio_override)` — enriquece cada adição e calcula totais
- Dicionários de alíquotas padrão por prefixo de NCM
- Cálculo encadeado: II → IPI (base = VA+II) → PIS/COFINS → ICMS

### `generator.py`
Gera o arquivo Excel com formatação profissional.
- `generate_excel(data, output_path)` — cria o workbook com 3 abas
- `_build_di_sheet()` — Aba 1: Extrato DI (layout RFB)
- `_build_adicoes_sheet()` — Aba 2: Tabela detalhada
- `_build_resumo_sheet()` — Aba 3: Dashboard tributário
- Paleta de cores azul institucional (RFB), linhas alternadas, bordas e totais em verde

### `app.py`
Servidor Flask com 4 rotas:
- `GET /` — página de upload
- `POST /converter` — processa e retorna Excel para download
- `POST /api/converter` — versão API que retorna JSON
- `GET /download/<id>` — download pelo ID gerado pela API
- `GET /health` — verificação de saúde

---

## 🛠️ Personalização

### Alterar alíquotas padrão
Em `processor.py`, edite os dicionários no topo do arquivo:
```python
_II_ALIQUOTAS_PADRAO = {
    "851220": 18.0,   # Faróis e lanternas
    "870829": 14.0,   # Peças diversas
    # Adicione seus NCMs aqui
}
```

### Alterar ICMS padrão
```python
_ICMS_PADRAO = 18.0  # Altere para o estado desejado
```

### Alterar porta do servidor
```bash
# Na linha de comando:
PORT=8080 python app.py

# Ou edite app.py:
app.run(port=8080)
```

---

## ❗ Tratamento de Erros

| Situação                         | Comportamento                               |
|----------------------------------|---------------------------------------------|
| Arquivo sem extensão suportada   | Mensagem de erro + redirect para home       |
| XML mal-formado                  | `ValueError` com detalhe do problema        |
| JSON inválido                    | `ValueError` com posição do erro            |
| Nenhuma adição encontrada        | Aviso amarelo + redirect para home          |
| Arquivo > 20MB                   | Erro 413 com mensagem amigável              |
| Campo ausente na adição          | Valor padrão (`0.0` ou `""`) sem exceção    |

---

## 📦 Dependências

| Biblioteca   | Versão mínima | Uso                          |
|--------------|---------------|------------------------------|
| flask        | 3.0.0         | Servidor web                 |
| pandas       | 2.0.0         | Leitura de Excel             |
| openpyxl     | 3.1.0         | Geração de Excel formatado   |
| werkzeug     | 3.0.0         | Upload seguro de arquivos    |
| lxml         | 4.9.0         | Parser XML alternativo       |

---

## 🧪 Testando sem interface

```python
from parser import parse_file
from processor import process_data
from generator import generate_excel

# Testar com XML exemplo
raw       = parse_file("exemplo_duimp.xml")
processed = process_data(raw, taxa_cambio_override=5.244)
generate_excel(processed, "outputs/meu_relatorio.xlsx")

print(processed["totais"])
```

---

## 📝 Notas Importantes

1. **ICMS é estimado** — o cálculo real pode variar por estado e regime tributário do adquirente
2. **II calculado automaticamente** usa alíquotas aproximadas da TEC; confirme sempre com o extrato oficial da DUIMP
3. Os arquivos de upload são **excluídos automaticamente** após o processamento
4. Os arquivos gerados em `/outputs` são limpos automaticamente após 2 horas
5. Para produção, use um servidor WSGI (Gunicorn/uWSGI) ao invés de `flask run`

---

## 🏗️ Melhorias Futuras

- [ ] Suporte a extrato PDF da DUIMP via OCR (pdfplumber)
- [ ] Cálculo de antidumping por NCM
- [ ] Exportação em formato PDF
- [ ] Banco de dados SQLite para histórico de processamentos
- [ ] Autenticação de usuário
- [ ] Integração com API SISCOMEX (quando disponível)
