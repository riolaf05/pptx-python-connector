# va-ppt-connector — Quickstart

Lambda AWS che compila un template PowerPoint (.pptx) con dati JSON e restituisce un link per il download.

## Struttura del progetto

```
va-ppt-connector/
├── examples/
│   ├── input.json                # JSON di esempio da inviare alla Lambda
│   └── report_template.pptx      # Template PPT con placeholder
├── lambda/
│   ├── handler.py                # Codice Lambda
│   └── requirements.txt          # Dipendenze Python
├── scripts/
│   └── create_template.py        # Script per (ri)generare il template PPT
├── terraform/
│   ├── main.tf                   # Infrastruttura AWS
│   ├── variables.tf              # Variabili Terraform
│   └── outputs.tf                # Output Terraform
└── QUICKSTART.md
```

---

## Come funzionano i placeholder nel PPT

Il template PPT contiene testo con **placeholder** nel formato `{{nome_campo}}`.
Esempio: una textbox contiene `{{titolo_report}}`.

Quando la Lambda riceve il JSON, cerca ogni occorrenza di `{{chiave}}` in tutte le textbox
di tutte le slide e la sostituisce con il valore corrispondente dal campo `placeholders`.

Per creare o modificare il template:

1. **A mano in PowerPoint**: apri il file `.pptx`, scrivi `{{nome_campo}}` dove vuoi il dato dinamico, salva.
2. **Con lo script Python**: modifica `scripts/create_template.py` e rigenera il file:
   ```bash
   source .venv/Scripts/activate   # Windows
   python scripts/create_template.py
   ```

I placeholder disponibili sono definiti dal JSON di input — puoi aggiungerne quanti ne vuoi,
basta che il nome nel PPT corrisponda alla chiave nel JSON.

---

## Prerequisiti

- Python 3.12+
- [Terraform](https://developer.hashicorp.com/terraform/install) >= 1.5
- AWS CLI configurato (`aws configure`)
- Un account AWS con permessi per creare S3, Lambda, IAM

---

## 1. Setup locale

```bash
# Clona il progetto e crea il virtual environment
cd va-ppt-connector
python -m venv .venv

# Attiva il venv
source .venv/Scripts/activate   # Windows (Git Bash)
# source .venv/bin/activate     # Linux/Mac

pip install python-pptx
```

## 2. (Opzionale) Rigenerare il template PPT

Se vuoi modificare le slide del template:

```bash
python scripts/create_template.py
# Output: examples/report_template.pptx
```

## 3. Deploy su AWS con Terraform

```bash
cd terraform
terraform init
terraform plan       # verifica cosa verrà creato
terraform apply      # conferma con "yes"
```

Terraform crea:
- **S3 bucket** — contiene template e output PPT
- **Lambda function** (`va-ppt-connector`) con layer per python-pptx
- **IAM role** — permessi S3 + CloudWatch Logs
- **Function URL** — endpoint HTTP pubblico per invocare la Lambda

Al termine vedrai gli output:

```
s3_bucket_name     = "va-ppt-connector-xxxx"
lambda_function_name = "va-ppt-connector"
lambda_function_url  = "https://xxxxx.lambda-url.eu-west-1.on.aws/"
```

## 4. Invocare la Lambda

### Via Function URL (HTTP POST)

```bash
curl -X POST "$(terraform output -raw lambda_function_url)" \
  -H "Content-Type: application/json" \
  -d @../examples/input.json
```

### Via AWS CLI

```bash
aws lambda invoke \
  --function-name va-ppt-connector \
  --payload file://../examples/input.json \
  --cli-binary-format raw-in-base64-out \
  response.json

cat response.json
```

### Risposta

```json
{
  "statusCode": 200,
  "body": "{\"message\": \"PPT compilato con successo\", \"output_s3_key\": \"output/report_compilato.pptx\", \"download_url\": \"https://s3.eu-west-1.amazonaws.com/...\"}"
}
```

Apri `download_url` nel browser per scaricare il PPT compilato. Il link scade dopo 1 ora (configurabile con `presigned_url_expiration`).

## 5. Cleanup

```bash
cd terraform
terraform destroy    # conferma con "yes"
```

---

## Personalizzazione

| Cosa | Come |
|---|---|
| Aggiungere campi | Aggiungi `{{nuovo_campo}}` nel PPT e la chiave corrispondente nel JSON |
| Cambiare regione | `terraform apply -var="aws_region=us-east-1"` |
| Durata link download | `terraform apply -var="presigned_url_expiration=7200"` |
| Template diversi | Carica più `.pptx` su S3 in `templates/` e specifica `template_s3_key` nel JSON |
