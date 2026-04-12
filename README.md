# va-ppt-connector

Servizio serverless su AWS che compila template PowerPoint (.pptx) con dati dinamici ricevuti via JSON e restituisce un link per il download del file generato.

## Panoramica

Il progetto risolve un problema comune: generare presentazioni PowerPoint personalizzate in modo automatico, partendo da un template con campi segnaposto e un set di dati.

### Architettura

```
                         ┌──────────────────────────────────────────────┐
                         │                  AWS                         │
                         │                                              │
  Client                 │   ┌──────────┐       ┌──────────────────┐   │
  (curl / app / browser) │   │          │  GET  │    S3 Bucket     │   │
         │               │   │  Lambda  │──────>│                  │   │
         │  POST JSON    │   │          │       │  templates/      │   │
         └──────────────>│   │  - scarica template                 │   │
                         │   │  - sostituisce placeholder          │   │
                         │   │  - salva output  │  output/         │   │
                         │   │  - genera URL    │                  │   │
                         │   │          │  PUT  │                  │   │
                         │   │          │──────>│                  │   │
                         │   └────┬─────┘       └──────────────────┘   │
                         │        │                                     │
                         └────────┼─────────────────────────────────────┘
                                  │
                                  ▼
                          Risposta JSON con
                          pre-signed URL di download
```

### Flusso

1. Il client invia una richiesta POST con un JSON che contiene il percorso del template e i valori da inserire
2. La Lambda scarica il template `.pptx` dal bucket S3
3. Cerca tutte le occorrenze di `{{nome_campo}}` nelle slide e le sostituisce con i valori ricevuti
4. Salva il file compilato su S3
5. Genera un pre-signed URL temporaneo e lo restituisce al client

## Sistema di placeholder

I template PowerPoint usano placeholder nel formato **`{{nome_campo}}`** inseriti come testo normale nelle textbox delle slide.

Esempio: una textbox con il testo `Fatturato: {{fatturato_totale}}` diventa `Fatturato: € 1.250.000` dopo la compilazione.

### Come creare un template

**Opzione 1 — In PowerPoint (consigliato per template complessi)**

1. Crea una presentazione in PowerPoint come faresti normalmente
2. Dove vuoi un dato dinamico, scrivi `{{nome_campo}}`
3. Salva come `.pptx` e caricalo su S3 nella cartella `templates/`

**Opzione 2 — Via script Python**

Modifica `scripts/create_template.py` per definire layout e placeholder via codice, poi esegui:

```bash
source .venv/Scripts/activate
python scripts/create_template.py
```

### Regole

- I nomi dei placeholder devono contenere solo lettere, numeri e underscore: `{{nome_campo_1}}`
- Lo stile del testo (font, colore, dimensione, grassetto) viene preservato dopo la sostituzione
- Puoi usare quanti placeholder vuoi, su quante slide vuoi
- Se un placeholder nel template non ha un valore corrispondente nel JSON, resta invariato

## Struttura del progetto

```
va-ppt-connector/
├── examples/
│   ├── input.json                 # JSON di input di esempio
│   └── report_template.pptx       # Template PPT con 5 slide di esempio
├── lambda/
│   ├── handler.py                 # Codice della Lambda
│   └── requirements.txt           # Dipendenze Python (python-pptx, boto3)
├── scripts/
│   └── create_template.py         # Generatore del template PPT di esempio
├── terraform/
│   ├── main.tf                    # S3, IAM, Lambda Layer, Lambda, Function URL
│   ├── variables.tf               # Variabili configurabili
│   └── outputs.tf                 # Output del deploy
├── .gitignore
├── QUICKSTART.md                  # Guida passo-passo per deploy e utilizzo
└── README.md
```

## Formato del JSON di input

```json
{
  "template_s3_key": "templates/report_template.pptx",
  "output_s3_key": "output/report_compilato.pptx",
  "placeholders": {
    "titolo_report": "Report Vendite Q1 2026",
    "autore": "Marco Rossi",
    "fatturato_totale": "€ 1.250.000"
  }
}
```

| Campo | Obbligatorio | Descrizione |
|---|---|---|
| `template_s3_key` | Si | Percorso del template `.pptx` nel bucket S3 |
| `output_s3_key` | No | Percorso di destinazione dell'output (default: `output/<request_id>.pptx`) |
| `placeholders` | No | Mappa chiave-valore dei campi da sostituire |

## Formato della risposta

```json
{
  "statusCode": 200,
  "body": {
    "message": "PPT compilato con successo",
    "output_s3_key": "output/report_compilato.pptx",
    "download_url": "https://s3.eu-west-1.amazonaws.com/..."
  }
}
```

Il `download_url` e' un pre-signed URL valido per 1 ora (configurabile).

## Componenti AWS

| Risorsa | Scopo |
|---|---|
| **S3 Bucket** | Archivia template e file compilati. Accesso pubblico bloccato. |
| **Lambda Function** | Esegue la compilazione. Runtime Python 3.12, 512 MB RAM, timeout 60s. |
| **Lambda Layer** | Contiene le dipendenze (`python-pptx`, `lxml`, `Pillow`). |
| **IAM Role** | Permessi per S3 (GetObject, PutObject) e CloudWatch Logs. |
| **Function URL** | Endpoint HTTP pubblico per invocare la Lambda senza API Gateway. |

## Tecnologie

- **Python 3.12** — runtime Lambda
- **python-pptx** — lettura e scrittura di file PowerPoint
- **Terraform** — Infrastructure as Code
- **AWS Lambda** — esecuzione serverless
- **AWS S3** — storage dei file

## Quickstart

Vedi [QUICKSTART.md](QUICKSTART.md) per le istruzioni dettagliate su prerequisiti, deploy e invocazione.

In breve:

```bash
# 1. Deploy infrastruttura
cd terraform
terraform init
terraform apply

# 2. Invoca la Lambda
curl -X POST "$(terraform output -raw lambda_function_url)" \
  -H "Content-Type: application/json" \
  -d @../examples/input.json

# 3. Apri il download_url dalla risposta per scaricare il PPT
```

## Configurazione

| Variabile Terraform | Default | Descrizione |
|---|---|---|
| `aws_region` | `eu-west-1` | Regione AWS |
| `project_name` | `va-ppt-connector` | Prefisso per le risorse |
| `presigned_url_expiration` | `3600` | Durata del link di download in secondi |

Esempio:

```bash
terraform apply \
  -var="aws_region=us-east-1" \
  -var="presigned_url_expiration=7200"
```
