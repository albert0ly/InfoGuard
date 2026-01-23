variable "environment" {
  description = "Environment name"
  type        = string
  default     = "dev"
}

variable "location" {
  description = "Azure region"
  type        = string
  default     = "eastus"
}

variable "project_name" {
  description = "Project name"
  type        = string
  default     = "InfoGuard"
}

variable "app_service_sku" {
  description = "App Service Plan SKU"
  type        = string
  default     = "B1"
}

variable "node_version" {
  description = "Node.js version"
  type        = string
  default     = "18-lts"
}

variable "allowed_origins" {
  description = "CORS allowed origins"
  type        = list(string)
  default     = ["https://outlook.office.com", "https://outlook.office365.com"]
}

variable "tags" {
  description = "Tags to apply"
  type        = map(string)
  default = {
    Project     = "InfoGuard"
    ManagedBy   = "Terraform"
  }
}