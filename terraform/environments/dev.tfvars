environment      = "dev"
location         = "eastus2"     # Change to eastus2
project_name     = "InfoGuard"
app_service_sku  = "B1"          # B1 will work now
node_version     = "18-lts"

allowed_origins = [
  "https://outlook.office.com",
  "https://outlook.office365.com",
  "https://localhost:3000"
]

tags = {
  Project     = "InfoGuard"
  Environment = "Development"
  ManagedBy   = "Terraform"
}