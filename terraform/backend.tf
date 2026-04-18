terraform {
  backend "s3" {
    bucket  = "va-ppt-connector-tfstate"
    key     = "va-ppt-connector/terraform.tfstate"
    region  = "eu-central-2"
    encrypt = true
  }
}
