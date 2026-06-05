terraform {
  required_version = ">= 1.5"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 5.0"
    }
  }
}

provider "aws" {
  region = var.aws_region

  default_tags {
    tags = {
      PROJECT = "VAF"
    }
  }
}

# -------------------------------------------------------------------
# S3 Bucket – contiene template e output
# -------------------------------------------------------------------
resource "aws_s3_bucket" "ppt_bucket" {
  bucket_prefix = "${var.project_name}-"
  force_destroy = true
}

resource "aws_s3_bucket_public_access_block" "ppt_bucket" {
  bucket                  = aws_s3_bucket.ppt_bucket.id
  block_public_acls       = true
  block_public_policy     = true
  ignore_public_acls      = true
  restrict_public_buckets = true
}

# Upload del template di esempio
resource "aws_s3_object" "ppt_template" {
  bucket       = aws_s3_bucket.ppt_bucket.id
  key          = "templates/report_template.pptx"
  source       = "${path.module}/../examples/report_template.pptx"
  content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  etag         = filemd5("${path.module}/../examples/report_template.pptx")
}

# Upload del template GTM Foundation Report
resource "aws_s3_object" "gtm_template" {
  bucket       = aws_s3_bucket.ppt_bucket.id
  key          = "templates/report_gtm_template.pptx"
  source       = "${path.module}/../examples/report_gtm_template.pptx"
  content_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  etag         = filemd5("${path.module}/../examples/report_gtm_template.pptx")
}

# -------------------------------------------------------------------
# IAM Role per la Lambda
# -------------------------------------------------------------------
data "aws_iam_policy_document" "lambda_assume" {
  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRole"]
    principals {
      type        = "Service"
      identifiers = ["lambda.amazonaws.com"]
    }
  }
}

resource "aws_iam_role" "lambda_role" {
  name               = "${var.project_name}-lambda-role"
  assume_role_policy = data.aws_iam_policy_document.lambda_assume.json
}

# Policy: CloudWatch Logs
resource "aws_iam_role_policy_attachment" "lambda_logs" {
  role       = aws_iam_role.lambda_role.name
  policy_arn = "arn:aws:iam::aws:policy/service-role/AWSLambdaBasicExecutionRole"
}

# Policy: S3 read/write sul bucket
data "aws_iam_policy_document" "lambda_s3" {
  statement {
    effect = "Allow"
    actions = [
      "s3:GetObject",
      "s3:PutObject",
    ]
    resources = ["${aws_s3_bucket.ppt_bucket.arn}/*"]
  }
}

resource "aws_iam_policy" "lambda_s3" {
  name   = "${var.project_name}-lambda-s3"
  policy = data.aws_iam_policy_document.lambda_s3.json
}

resource "aws_iam_role_policy_attachment" "lambda_s3" {
  role       = aws_iam_role.lambda_role.name
  policy_arn = aws_iam_policy.lambda_s3.arn
}

# Policy: SES – invio email di alert
data "aws_iam_policy_document" "lambda_ses" {
  statement {
    effect    = "Allow"
    actions   = ["ses:SendEmail", "ses:SendRawEmail"]
    resources = ["*"]
  }
}

resource "aws_iam_policy" "lambda_ses" {
  name   = "${var.project_name}-lambda-ses"
  policy = data.aws_iam_policy_document.lambda_ses.json
}

resource "aws_iam_role_policy_attachment" "lambda_ses" {
  role       = aws_iam_role.lambda_role.name
  policy_arn = aws_iam_policy.lambda_ses.arn
}

# -------------------------------------------------------------------
# SES – verifica email di alert
# -------------------------------------------------------------------
resource "aws_ses_email_identity" "alert" {
  email = var.alert_email
}

# -------------------------------------------------------------------
# Lambda Layer – dipendenze python-pptx
# -------------------------------------------------------------------
resource "null_resource" "build_layer" {
  triggers = {
    requirements = filemd5("${path.module}/../lambda/requirements.txt")
  }

  provisioner "local-exec" {
    interpreter = ["powershell", "-Command"]
    command     = <<-EOT
      if (Test-Path "${path.module}/../build/layer") { Remove-Item -Recurse -Force "${path.module}/../build/layer" }
      New-Item -ItemType Directory -Force -Path "${path.module}/../build/layer/python" | Out-Null
      pip install -r "${path.module}/../lambda/requirements.txt" `
        -t "${path.module}/../build/layer/python" `
        --platform manylinux2014_x86_64 `
        --implementation cp `
        --python-version 3.12 `
        --only-binary=:all: `
        --no-deps
      if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
      pip install python-pptx==1.0.2 lxml typing-extensions Pillow XlsxWriter `
        -t "${path.module}/../build/layer/python" `
        --platform manylinux2014_x86_64 `
        --implementation cp `
        --python-version 3.12 `
        --only-binary=:all:
      if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
    EOT
  }
}

data "archive_file" "layer_zip" {
  type        = "zip"
  source_dir  = "${path.module}/../build/layer"
  output_path = "${path.module}/../build/layer.zip"

  depends_on = [null_resource.build_layer]
}

resource "aws_lambda_layer_version" "deps" {
  filename            = data.archive_file.layer_zip.output_path
  source_code_hash    = data.archive_file.layer_zip.output_base64sha256
  layer_name          = "${var.project_name}-deps"
  compatible_runtimes = ["python3.12"]
}

# -------------------------------------------------------------------
# Lambda Function
# -------------------------------------------------------------------
data "archive_file" "lambda_zip" {
  type        = "zip"
  source_dir  = "${path.module}/../lambda"
  output_path = "${path.module}/../build/lambda.zip"
}

resource "aws_lambda_function" "ppt_compiler" {
  function_name    = "${var.project_name}"
  role             = aws_iam_role.lambda_role.arn
  handler          = "handler.lambda_handler"
  runtime          = "python3.12"
  timeout          = 60
  memory_size      = 512
  filename         = data.archive_file.lambda_zip.output_path
  source_code_hash = data.archive_file.lambda_zip.output_base64sha256
  layers           = [aws_lambda_layer_version.deps.arn]

  environment {
    variables = {
      S3_BUCKET                = aws_s3_bucket.ppt_bucket.id
      PRESIGNED_URL_EXPIRATION = tostring(var.presigned_url_expiration)
      ALERT_EMAIL              = var.alert_email
      ALERT_FROM_EMAIL         = var.alert_email
      API_KEY                  = var.api_key
    }
  }
}

# -------------------------------------------------------------------
# Lambda Function URL (accesso pubblico senza API Gateway)
# -------------------------------------------------------------------
resource "aws_lambda_function_url" "ppt_compiler" {
  function_name      = aws_lambda_function.ppt_compiler.function_name
  authorization_type = "NONE"
}

# -------------------------------------------------------------------
# Lambda Function – GTM Foundation Report compiler
# Richiede l'header x-api-key per l'autenticazione.
# -------------------------------------------------------------------
resource "aws_lambda_function" "gtm_compiler" {
  function_name    = "${var.project_name}-gtm"
  role             = aws_iam_role.lambda_role.arn
  handler          = "gtm_handler.lambda_handler"
  runtime          = "python3.12"
  timeout          = 60
  memory_size      = 512
  filename         = data.archive_file.lambda_zip.output_path
  source_code_hash = data.archive_file.lambda_zip.output_base64sha256
  layers           = [aws_lambda_layer_version.deps.arn]

  environment {
    variables = {
      S3_BUCKET                = aws_s3_bucket.ppt_bucket.id
      GTM_TEMPLATE_KEY         = "templates/report_gtm_template.pptx"
      PRESIGNED_URL_EXPIRATION = tostring(var.presigned_url_expiration)
      ALERT_EMAIL              = var.alert_email
      ALERT_FROM_EMAIL         = var.alert_email
      API_KEY                  = var.gtm_api_key
    }
  }
}

resource "aws_lambda_function_url" "gtm_compiler" {
  function_name      = aws_lambda_function.gtm_compiler.function_name
  authorization_type = "NONE"
}
