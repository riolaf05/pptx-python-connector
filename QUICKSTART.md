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
│   ├── backend.tf                # Backend S3 per lo state Terraform (Zurigo)
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
   ```powershell
   .venv\Scripts\Activate.ps1
   python scripts/create_template.py
   ```

I placeholder disponibili sono definiti dal JSON di input — puoi aggiungerne quanti ne vuoi,
basta che il nome nel PPT corrisponda alla chiave nel JSON.

> **Nota encoding**: nel JSON usa sempre escape Unicode per caratteri speciali (es. `\u20ac` per `€`).
> Il file `examples/input.json` è già configurato correttamente.

---

## Prerequisiti

- Python 3.12+
- [Terraform](https://developer.hashicorp.com/terraform/install) >= 1.5
- AWS CLI configurato (`aws configure`)
- Un account AWS con permessi per creare S3, Lambda, IAM

---

## 1. Setup locale

```powershell
# Crea il virtual environment
cd va-ppt-connector
python -m venv .venv

# Attiva il venv (PowerShell)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.venv\Scripts\Activate.ps1

pip install python-pptx
```

## 2. (Opzionale) Rigenerare il template PPT

Se vuoi modificare le slide del template:

```powershell
python scripts/create_template.py
# Output: examples/report_template.pptx
```

## 3. Crea il bucket S3 per lo Terraform state (una tantum)

Il backend è configurato su **Zurigo** (`eu-central-2`). Esegui questi comandi **solo la prima volta**:

```powershell
# Crea il bucket
aws s3api create-bucket `
  --bucket va-ppt-connector-tfstate `
  --region eu-central-2 `
  --create-bucket-configuration LocationConstraint=eu-central-2

# Versioning
aws s3api put-bucket-versioning `
  --bucket va-ppt-connector-tfstate `
  --versioning-configuration Status=Enabled

# Crittografia
aws s3api put-bucket-encryption `
  --bucket va-ppt-connector-tfstate `
  --server-side-encryption-configuration '{\"Rules\":[{\"ApplyServerSideEncryptionByDefault\":{\"SSEAlgorithm\":\"AES256\"},\"BucketKeyEnabled\":true}]}'

# Blocco accesso pubblico
aws s3api put-public-access-block `
  --bucket va-ppt-connector-tfstate `
  --public-access-block-configuration "BlockPublicAcls=true,IgnorePublicAcls=true,BlockPublicPolicy=true,RestrictPublicBuckets=true"

# Tag progetto
aws s3api put-bucket-tagging `
  --bucket va-ppt-connector-tfstate `
  --tagging 'TagSet=[{Key=PROJECT,Value=VAF}]'
```

## 4. Deploy su AWS con Terraform

```powershell
cd terraform

# Prima inizializzazione (oppure -migrate-state se esiste uno state locale)
terraform init

terraform plan       # verifica cosa verrà creato
terraform apply      # conferma con "yes"
```

In caso di errore sul `null_resource` (build del layer Python), forza la ri-esecuzione:

```powershell
terraform apply -replace="null_resource.build_layer"
```

Terraform crea:
- **S3 bucket** — contiene template e output PPT (tag `PROJECT: VAF` automatico su tutte le risorse)
- **Lambda function** (`va-ppt-connector`) con layer per python-pptx
- **IAM role** — permessi S3 + CloudWatch Logs
- **Function URL** — endpoint HTTP pubblico per invocare la Lambda

Al termine vedrai gli output:

```
s3_bucket_name       = "va-ppt-connector-xxxx"
lambda_function_name = "va-ppt-connector"
lambda_function_url  = "https://xxxxx.lambda-url.eu-west-1.on.aws/"
```

## 5. Caricare un nuovo template su S3

```powershell
$BUCKET = terraform output -raw s3_bucket_name

aws s3 cp "..\examples\report_template.pptx" "s3://$BUCKET/templates/report_template.pptx"
```

## 6. Test rapido — compilare il PPT e scaricare il risultato

```powershell
cd ..   # torna nella root del progetto

# Invoca la Lambda con il JSON di esempio
aws lambda invoke `
  --function-name va-ppt-connector `
  --region eu-west-1 `
  --payload file://examples/input.json `
  --cli-binary-format raw-in-base64-out `
  response.json

# Estrai il pre-signed URL e scarica il file
$url = (Get-Content response.json | ConvertFrom-Json | `
  Select-Object -ExpandProperty body | ConvertFrom-Json).download_url

Invoke-WebRequest -Uri $url -OutFile ".\report_compilato.pptx"
```

Il file `report_compilato.pptx` viene salvato nella root del progetto.

### Risposta Lambda

```json
{
  "statusCode": 200,
  "body": "{\"message\": \"PPT compilato con successo\", \"output_s3_key\": \"output/report_compilato.pptx\", \"download_url\": \"https://...\"}"
}
```

Il pre-signed URL scade dopo **1 ora** (configurabile con la variabile `presigned_url_expiration`).

## 7. Cleanup

```powershell
cd terraform
terraform destroy    # conferma con "yes"
```

---

## Personalizzazione

| Cosa | Come |
|---|---|
| Aggiungere campi | Aggiungi `{{nuovo_campo}}` nel PPT e la chiave corrispondente nel JSON |
| Cambiare regione | `terraform apply -var="aws_region=eu-central-1"` |
| Durata link download | `terraform apply -var="presigned_url_expiration=7200"` |
| Template diversi | Carica più `.pptx` su S3 in `templates/` e specifica `template_s3_key` nel JSON |
