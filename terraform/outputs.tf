output "s3_bucket_name" {
  description = "Nome del bucket S3"
  value       = aws_s3_bucket.ppt_bucket.id
}

output "lambda_function_name" {
  description = "Nome della Lambda"
  value       = aws_lambda_function.ppt_compiler.function_name
}

output "lambda_function_url" {
  description = "URL pubblico della Lambda (Function URL)"
  value       = aws_lambda_function_url.ppt_compiler.function_url
}

output "gtm_lambda_function_name" {
  description = "Nome della Lambda GTM"
  value       = aws_lambda_function.gtm_compiler.function_name
}

output "gtm_lambda_function_url" {
  description = "URL pubblico della Lambda GTM (Function URL)"
  value       = aws_lambda_function_url.gtm_compiler.function_url
}
