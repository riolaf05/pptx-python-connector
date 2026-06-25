# va-ppt-connector — Guida per LLM: come chiamare gli endpoint

Questo documento descrive come invocare le Lambda di va-ppt-connector per generare un file PowerPoint compilato e scaricarlo tramite pre-signed URL.

---

## Contesto

Il sistema espone **due Lambda** via HTTP (AWS Lambda Function URL):

| Lambda | Scopo | Handler |
|---|---|---|
| `va-ppt-connector` | Report analitico generico (SWOT, findings) | `handler.lambda_handler` |
| `va-ppt-connector-gtm` | GTM Foundation Report (ICP, Persona, USP, ecc.) | `gtm_handler.lambda_handler` |

Entrambe richiedono autenticazione via header `x-api-key`.

---

## Flusso generale

```
1. Chiama l'endpoint HTTP (POST) con JSON + x-api-key
2. La Lambda scarica il template da S3, sostituisce i placeholder, carica il file compilato su S3
3. La risposta contiene un `download_url` (AWS S3 pre-signed URL, valido 1 ora)
4. Fai una GET al `download_url` per scaricare il file .pptx
```

---

## Endpoint 1 — Report Analitico (`va-ppt-connector`)

### Richiesta

```
POST <LAMBDA_URL>
Content-Type: application/json
x-api-key: <API_KEY>
```

### Body JSON (formato n8n / strutturato)

```json
[
  {
    "output": {
      "strengths_synthesis_items": ["Punto di forza 1", "Punto di forza 2"],
      "weaknesses_synthesis_items": ["Debolezza 1", "Debolezza 2"],
      "opportunities_synthesis_items": ["Opportunità 1"],
      "threats_synthesis_items": ["Minaccia 1"],
      "sales_so_finding": ["Finding SO vendite 1", "Finding SO vendite 2"],
      "sales_tb_finding": ["Finding TB vendite 1"],
      "marketing_so_finding": ["Finding SO marketing 1"],
      "marketing_tb_finding": ["Finding TB marketing 1"],
      "product_so_finding": ["Finding SO prodotto 1"],
      "product_tb_finding": ["Finding TB prodotto 1"],
      "strategic_so_finding": ["Finding SO strategico 1"],
      "strategic_tb_finding": ["Finding TB strategico 1"]
    },
    "customer_name": "Nome Cliente",
    "document_date": "05.06.2026",
    "document_version": "1.0",
    "assessment_period": "Q1 2026",
    "total_document_reviewed": "42",
    "total_interviews": "12",
    "template_s3_key": "templates/report_template.pptx",
    "output_s3_key": "output/report_cliente.pptx"
  }
]
```

**Note sui campi array:**
- `strengths_synthesis_items` → placeholder `{{swot_strength_1}}` … `{{swot_strength_4}}` (max 4)
- `weaknesses_synthesis_items` → `{{swot_weakness_1}}` … `{{swot_weakness_4}}` (max 4)
- `opportunities_synthesis_items` → `{{swot_opportunity_1}}` … `{{swot_opportunity_4}}` (max 4)
- `threats_synthesis_items` → `{{swot_threat_1}}` … `{{swot_threat_4}}` (max 4)
- `sales_so_finding` → `{{sales_so_finding_1}}` … `{{sales_so_finding_9}}` (max 9)
- `sales_tb_finding` → `{{sales_tb_finding_1}}` … `{{sales_tb_finding_3}}` (max 3)
- Idem per `marketing_*`, `product_*`, `strategic_*`

### Body JSON alternativo (formato diretto con placeholder flat)

```json
{
  "template_s3_key": "templates/report_template.pptx",
  "output_s3_key": "output/report_cliente.pptx",
  "placeholders": {
    "customer_name": "Nome Cliente",
    "document_date": "05.06.2026",
    "document_version": "1.0",
    "assessment_period": "Q1 2026",
    "total_document_reviewed": "42",
    "total_interviews": "12",
    "swot_strength_1": "Punto di forza 1",
    "swot_strength_2": "Punto di forza 2",
    "swot_weakness_1": "Debolezza 1",
    "sales_so_finding_1": "Finding SO vendite 1"
  }
}
```

---

## Endpoint 2 — GTM Foundation Report (`va-ppt-connector-gtm`)

### Richiesta

```
POST <GTM_LAMBDA_URL>
Content-Type: application/json
x-api-key: <GTM_API_KEY>
```

### Body JSON completo

```json
{
  "customer_name": "Nome Cliente",
  "document_date": "DD.MM.YYYY",
  "document_version": "1.0",
  "template_s3_key": "templates/report_gtm_template.pptx",
  "output_s3_key": "output/gtm_report_cliente.pptx",

  "icp": {
    "industry": "Settore di appartenenza",
    "geography": "Area geografica target",
    "company_size": "Dimensione azienda",
    "buying_group": "Gruppo di acquisto (ruoli)",
    "maturity_level": "Livello di maturità digitale",
    "custom_field": "Campo personalizzabile"
  },

  "persona": {
    "role_name": "Titolo del ruolo",
    "name": "Nome Cognome",
    "role": "Ruolo completo",
    "age": "40",
    "education": "Titolo di studio",
    "location": "Città, Paese",
    "industry": "Settore",
    "channels": "LinkedIn | Email | Events",
    "responsibilities": [
      "Responsabilità 1",
      "Responsabilità 2",
      "Responsabilità 3"
    ],
    "objective_1": {
      "title": "Titolo Obiettivo 1",
      "description": "Descrizione dell'obiettivo",
      "kpi1": "KPI 1",
      "kpi2": "KPI 2",
      "kpi3": "KPI 3",
      "pain_point_1": "Pain point principale",
      "pain_point_2": "Pain point secondario",
      "winning_trigger": "Trigger di acquisto"
    },
    "objective_2": { "...": "stessa struttura di objective_1" },
    "objective_3": { "...": "stessa struttura di objective_1" },
    "objective_4": { "...": "stessa struttura di objective_1" }
  },

  "value_proposition": {
    "what_is": "Cos'è il prodotto/servizio",
    "what_does": "Cosa fa il prodotto/servizio",
    "how_works": "Come funziona (in termini semplici)",
    "value_to_customer": "Valore per il cliente"
  },

  "slogans": [
    "Slogan 1",
    "Slogan 2",
    "Slogan 3"
  ],

  "customer_mental_model": {
    "persona": "Nome del target persona",
    "objective": "Obiettivo principale del cliente",
    "driving_factors": [
      "Fattore guida 1",
      "Fattore guida 2"
    ],
    "pain_points": [
      "Pain point 1",
      "Pain point 2"
    ],
    "use_cases": [
      "Caso d'uso 1",
      "Caso d'uso 2"
    ]
  },

  "usps": [
    {
      "title": "Titolo USP 1",
      "body": ["Punto 1", "Punto 2", "Punto 3"]
    },
    {
      "title": "Titolo USP 2",
      "body": ["Punto 1", "Punto 2", "Punto 3"]
    },
    {
      "title": "Titolo USP 3",
      "body": ["Punto 1", "Punto 2", "Punto 3"]
    }
  ],

  "commercial_insight": {
    "objective": "Obiettivo di business (nodo 1)",
    "pain_point": "Pain point (nodo 2)",
    "use_case": "Use case (nodo 3)",
    "usp": "USP (nodo 4)",
    "title": "Titolo dell'insight",
    "narratives": [
      "Frase 1: Lavoriamo con [persona] che vogliono [obiettivo] ma affrontano [problema].",
      "Frase 2: Normalmente provano [approccio comune] ma il problema si risolve con [use case].",
      "Frase 3: Siamo nella posizione unica per farlo grazie a [USP]."
    ]
  }
}
```

---

## Risposta (identica per entrambi gli endpoint)

### Successo — HTTP 200

```json
{
  "message": "GTM Report compilato con successo",
  "output_s3_key": "output/gtm_report_cliente.pptx",
  "download_url": "https://s3.eu-west-1.amazonaws.com/va-ppt-connector-xxxx/output/gtm_report_cliente.pptx?X-Amz-Algorithm=..."
}
```

### Errore autenticazione — HTTP 401

```json
{
  "error": "Unauthorized: x-api-key mancante o non valida"
}
```

### Errore interno — HTTP 500

```json
{
  "error": "Traceback (most recent call last): ..."
}
```

---

## Come scaricare il file

Il campo `download_url` è un **AWS S3 pre-signed URL**:
- È una URL HTTPS pubblica temporanea, **non richiede autenticazione aggiuntiva**
- Scade dopo **1 ora** dalla generazione
- Va usata con una semplice GET HTTP

```
GET <download_url>
→ risposta: file binario .pptx
   Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation
```

---

## Regole e limiti

| Parametro | Valore |
|---|---|
| Metodo HTTP | `POST` |
| Content-Type richiesto | `application/json` |
| Header autenticazione | `x-api-key: <chiave>` |
| Scadenza pre-signed URL | 3600 secondi (1 ora) |
| Timeout Lambda | 60 secondi |
| Formato template | `.pptx` (su S3 in `templates/`) |
| Formato output | `.pptx` (su S3 in `output/`) |
| `output_s3_key` | Opzionale — se omesso viene generato automaticamente |
| `template_s3_key` | Opzionale — usa il default se omesso |

---

## Campi opzionali e default

| Campo | Default se omesso |
|---|---|
| `template_s3_key` (GTM) | `templates/report_gtm_template.pptx` |
| `template_s3_key` (Report) | `templates/report_template.pptx` |
| `output_s3_key` | `output/<request-id>.pptx` (generato automaticamente) |
| `document_version` | `1.0` |

---

## Esempio minimo GTM (campi obbligatori soltanto)

```json
{
  "customer_name": "Acme Srl",
  "document_date": "05.06.2026",
  "icp": { "industry": "SaaS B2B" },
  "value_proposition": { "what_is": "Una piattaforma CRM avanzata" },
  "usps": [{ "title": "Veloce", "body": ["Setup in 1 giorno"] }],
  "commercial_insight": {
    "title": "L'insight principale",
    "narratives": ["Frase 1", "Frase 2", "Frase 3"]
  }
}
```

I campi non forniti vengono lasciati vuoti nel documento finale.
