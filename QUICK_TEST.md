# va-ppt-connector — Test rapido (PowerShell)

Guida per testare l'invocazione della Lambda e scaricare il PPT compilato.
Assicurati di aver già eseguito il deploy con `terraform apply`.

---

## 0. Variabili di sessione

```powershell
cd C:\Users\r.laface\Desktop\Codice\va-ppt-connector\terraform

$BUCKET     = terraform output -raw s3_bucket_name
$LAMBDA_URL = terraform output -raw lambda_function_url

Write-Host "Bucket : $BUCKET"
Write-Host "Lambda : $LAMBDA_URL"
```

---

## 1. Caricare un template su S3

```powershell
aws s3 cp "..\examples\report_template.pptx" `
  "s3://$BUCKET/templates/report_template.pptx"
```

---

## 2. Compilare il PPT — via AWS CLI (consigliato su Windows)

```powershell
cd C:\Users\r.laface\Desktop\Codice\va-ppt-connector

aws lambda invoke `
  --function-name va-ppt-connector `
  --region eu-west-1 `
  --payload file://examples/input.json `
  --cli-binary-format raw-in-base64-out `
  response.json

# Visualizza la risposta
Get-Content response.json | ConvertFrom-Json | Select-Object -ExpandProperty body | ConvertFrom-Json
```

---

## 3. Scaricare il PPT compilato

```powershell
$url = (Get-Content response.json | ConvertFrom-Json | `
  Select-Object -ExpandProperty body | ConvertFrom-Json).download_url

Invoke-WebRequest -Uri $url -OutFile ".\report_compilato.pptx"

Write-Host "File salvato: report_compilato.pptx"
```

Il pre-signed URL scade dopo **1 ora**.

---

## 4. Compilare con dati personalizzati — via Function URL

```powershell
$body = @{
    template_s3_key = "templates/report_template.pptx"
    output_s3_key   = "output/report_custom.pptx"
    placeholders    = @{
        customer_name            = "Acme S.p.A."
        document_date            = "29 Aprile 2026"
        document_version         = "1.0"
        assessment_period        = "Q1 2026"
        total_document_reviewed  = "42"
        total_interviews         = "12"
        swot_strength_1          = "Brand riconosciuto sul mercato nazionale"
        swot_strength_2          = "Team tecnico altamente specializzato"
        swot_strength_3          = "Portafoglio prodotti diversificato"
        swot_strength_4          = "Solida base clienti enterprise"
        swot_weakness_1          = "Processi interni non ancora digitalizzati"
        swot_weakness_2          = "Dipendenza da pochi fornitori chiave"
        swot_weakness_3          = "Time-to-market elevato"
        swot_weakness_4          = "Scarsa presenza sui canali digitali"
        swot_opportunity_1       = "Crescita del mercato cloud in EMEA"
        swot_opportunity_2       = "Nuovi segmenti SMB non ancora presidiati"
        swot_opportunity_3       = "Partnership con system integrator"
        swot_opportunity_4       = "Espansione geografica nel mercato DACH"
        swot_threat_1            = "Ingresso di nuovi competitor low-cost"
        swot_threat_2            = "Evoluzione normativa GDPR e AI Act"
        swot_threat_3            = "Pressione sui margini da parte dei clienti"
        swot_threat_4            = "Volatilit$([char]0x00E0) dei costi delle materie prime"
        # ... aggiungi tutti i finding (vedi examples/input.json per la lista completa)
    }
} | ConvertTo-Json -Depth 5 -EscapeHandling EscapeNonAscii

$response = Invoke-RestMethod `
  -Uri $LAMBDA_URL `
  -Method POST `
  -ContentType "application/json" `
  -Body $body

# Scarica il risultato
Invoke-WebRequest -Uri $response.download_url -OutFile ".\report_custom.pptx"
Write-Host "File salvato: report_custom.pptx"
```

> **Suggerimento**: per invocazioni complete usa sempre il metodo AWS CLI (sezione 2)
> con `file://examples/input.json` — è più semplice e non ha problemi di encoding.

---

## 5. Verificare i log CloudWatch

```powershell
aws logs filter-log-events `
  --log-group-name /aws/lambda/va-ppt-connector `
  --region eu-west-1 `
  --limit 20 `
  --query "events[*].message" `
  --output text
```

---

## Tutto in un unico script

```powershell
cd C:\Users\r.laface\Desktop\Codice\va-ppt-connector

# 1. Upload template
$BUCKET = terraform -chdir=".\terraform" output -raw s3_bucket_name
aws s3 cp ".\examples\report_template.pptx" "s3://$BUCKET/templates/report_template.pptx"

# 2. Invoca Lambda
aws lambda invoke `
  --function-name va-ppt-connector `
  --region eu-west-1 `
  --payload file://examples/input.json `
  --cli-binary-format raw-in-base64-out `
  response.json

# 3. Scarica il PPT
$url = (Get-Content response.json | ConvertFrom-Json | `
  Select-Object -ExpandProperty body | ConvertFrom-Json).download_url
Invoke-WebRequest -Uri $url -OutFile ".\report_compilato.pptx"
Write-Host "Fatto: report_compilato.pptx"
```
