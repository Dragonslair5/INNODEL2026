# INNODEL 2026 — Bibliometric Analysis of Generative AI in Online and Distance Education

Analysis code for the paper **"Mapping an Emerging Research Field: A Bibliometric Analysis of Generative AI in Online and Distance Education"** (2026 INNODEL, UNIVESP — Universidade Virtual do Estado de São Paulo).

Authors: Marina Delege, Mauricio G. Palma, Tabata Rejane Bertazzo, Tamires Cristina Bonani Conti.

The script reads a consolidated Scopus / Web of Science / ScienceDirect export (one Excel sheet) and produces all charts, networks and LaTeX tables used in the paper.



## What it produces

| # | Analysis | Outputs |
|---|----------|---------|
| – | Dataset summary and data-coverage table | console summary, `data_coverage.tex` |
| 1 | Publications per year | bar chart |
| 2 | Top journals by number of publications | bar chart, dual-metric chart |
| 3 | Top journals by total citations | bar chart, dual-metric chart, citations-per-paper ranking, publications-vs-citations scatter (with zoom panel) |
| 4 | Most-cited papers | CSV, LaTeX table with journal rank and citations-per-publication ratio |
| 5 | Average citations per year | chart and LaTeX table (raw and normalized by years since publication) |
| 6 | Most frequent keywords | bar chart, word cloud |
| 7 | Keyword co-occurrence network | SVG + PDF figure, interactive HTML, node-statistics LaTeX table |
| 8 | Author co-authorship network | SVG + PDF figure, interactive HTML, node-statistics LaTeX table |
| 9 | Lotka's Law check | chart with the classical (c = 2) and empirically fitted curves, fit statistics |

Every ranking states the total it was drawn from (for example, "top 15 of 215 journals") in its title and caption.

## Installation

Python 3.9 or later.

```bash
pip install pandas numpy openpyxl matplotlib networkx scipy wordcloud pyvis adjustText
```

`wordcloud`, `pyvis` and `adjustText` are only needed for the word cloud, the interactive HTML networks and the non-overlapping labels in the scatter plot, respectively.

## Usage

```bash
python bibliometric_analysis.py path/to/dataset.xlsx --sheet "Sheet1" --top_n 15
```

| Argument | Default | Description |
|----------|---------|-------------|
| `excel_path` | – | Path to the `.xlsx` file |
| `--sheet` | first sheet | Sheet name or index |
| `--top_n` | 15 | Number of items in the "top N" rankings |

All results are written to `./output/`.

Usage example:
```bash
python bibliometric_analysis.py ./Resultados_unicos.xlsx --sheet Resultados_unicos --top_n 15
```


### Configuration

Constants at the top of the script control the defaults:

```python
TOP_N = 15                 # items in "top N" charts and tables
KW_MIN_COOCCUR = 2         # minimum co-occurrences to keep a keyword edge
AUTHOR_MIN_COOCCUR = 2     # minimum joint publications to keep an author edge
OUTPUT_DIR = "output"
```

The column names the script expects are defined in the same block (`COL_TITLE`, `COL_YEAR`, and so on). Edit them if your export uses different headers.

## Expected input

One Excel sheet, one row per document, with these columns (Portuguese headers):

| Column | Content |
|--------|---------|
| `Titulo` | Document title |
| `Ano` | Publication year |
| `Periodico` | Journal |
| `Autores` | Authors, separated by `;` |
| `Citacoes_Scopus`, `Citacoes_WoS` | Citation counts (Scopus is used when present, WoS as fallback) |
| `Palavras_chave_autor` | Author keywords |
| `Palavras_chave_indexadas` | Indexed keywords |

The coverage table additionally reports these optional columns when present: `Afiliacoes`, `Autores_com_afiliacao`, `Idioma`, `Editora`, `Fasciculo`, `Financiamento`.

## Method notes

- **Corpus.** Records from Web of Science (25), Scopus (102) and ScienceDirect (636) were merged by DOI, giving 735 documents after removing 28 duplicates (search run on 7 September 2026). The ScienceDirect query was reduced because of that database's limit of eight boolean operators.
- **Citation coverage.** Databases export different metadata, and only 14.7% of records carry citation counts. All citation-based results use only the records that have them.
- **2026 and 2027.** 2026 is a partial year, as the query was performed on September 7, 2026. Documents labelled 2027 are issues already indexed under 2027 volumes.
- **Co-occurrence networks.** Nodes are the most frequent keywords (or the most prolific authors); an edge links two nodes that appear together in at least the minimum number of documents. Edge thickness is proportional to the co-occurrence count.
- **Lotka's Law.** Observed author productivity is compared with the classical inverse-square law and with an exponent fitted to the data by log–log linear regression. Goodness of fit uses a grouped-frequency Kolmogorov–Smirnov-type statistic against the critical value 1.36/√N (Pao, 1985). The usual one-sample K–S test is not used because the data are heavily tied (most published only once).

## Repository layout

```
.
├── README.md
├── INNODEL_2026_MMTT_PRESENTATION.pdf # presentation slides
├── Resultados_unicos.xlsx             # our consolidated database
├── bibliometric_analysis.py           # analysis script
└── output/                            # output of analysis script
```


## Contact

Mauricio G. Palma — mauricio.palma@univesp.br