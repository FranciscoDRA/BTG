# BTG PDF → Excel

Convierte los extractos de cuenta en PDF de **BTG Pactual Uruguay (iBanca, "Últimos movimientos")** a un Excel con las columnas:

| Fecha | Concepto | Débito | Crédito |
|---|---|---|---|
| 05/08/26 | Pago Préstamo 199558 | 918881.00 | 0.00 |
| 05/08/26 | Descuento de documentos 291068 | 0.00 | 918881.29 |

- El concepto incluye la referencia o el cheque.
- Los débitos van en positivo.
- Cada fila se valida contra el saldo del PDF (saldo anterior − débito + crédito), y el programa avisa si algo no cierra.

## Usarlo online (sin instalar nada)

👉 **https://franciscodra.github.io/BTG/**

Arrastrás el PDF y descargás el Excel. Todo se procesa en tu navegador: el PDF no se sube a ningún servidor.

## Versión Python (consola / API)

### Instalación

Requiere Python 3.9 o superior.

```bash
git clone https://github.com/FranciscoDRA/BTG.git
cd BTG
pip install -r requirements.txt
```

### Uso por consola

```bash
python btg_a_excel.py extracto_julio.pdf extracto_agosto.pdf
```

Genera `extracto_julio_extraido.xlsx` y `extracto_agosto_extraido.xlsx` al lado de cada PDF.

### Uso como API

```bash
uvicorn btg_a_excel:app --reload
```

Abrí <http://localhost:8000/docs>, entrá a `POST /convertir`, tocá **Try it out**, subí el PDF y descargá el Excel.

Con curl:

```bash
curl -F "archivo=@extracto.pdf" http://localhost:8000/convertir -o extracto.xlsx
```

## Archivos

| Archivo | Para qué |
|---|---|
| `index.html` + `parser.js` | Página web (GitHub Pages) |
| `btg_a_excel.py` | Versión Python: consola y API FastAPI |

## Limitaciones

- Solo está probado con extractos de BTG en pesos (UYU) exportados desde iBanca.
- Si la columna Saldo sale recortada en el PDF (pasa al imprimir desde iBanca), igual funciona: los importes son exactos y el saldo se verifica sin centavos.
- Si el PDF es una imagen escaneada sin texto, no funciona porque no hace OCR.

> ⚠️ Los extractos tienen datos bancarios reales. El `.gitignore` ya excluye `*.pdf` y `*.xlsx`: no los subas al repo.
