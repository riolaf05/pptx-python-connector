variable "aws_region" {
  description = "AWS region"
  type        = string
  default     = "eu-west-1"
}

variable "project_name" {
  description = "Nome del progetto, usato come prefisso per le risorse"
  type        = string
  default     = "va-ppt-connector"
}

variable "presigned_url_expiration" {
  description = "Durata in secondi del pre-signed URL (default 1h)"
  type        = number
  default     = 3600
}

variable "alert_email" {
  description = "Indirizzo email per gli alert di errore Lambda (deve essere verificato in SES)"
  type        = string
  default     = "r.laface@value-accelerator.io"
}
