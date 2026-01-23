output "backend_url" {
  description = "Backend URL"
  value       = "https://${azurerm_linux_web_app.backend.default_hostname}"
}

output "backend_name" {
  description = "App Service name"
  value       = azurerm_linux_web_app.backend.name
}

output "resource_group_name" {
  description = "Resource group name"
  value       = azurerm_resource_group.main.name
}

output "storage_account_name" {
  description = "Storage account name"
  value       = azurerm_storage_account.addin.name
}

output "key_vault_name" {
  description = "Key Vault name"
  value       = azurerm_key_vault.main.name
}

output "tenant_id" {
  description = "Tenant ID"
  value       = data.azurerm_client_config.current.tenant_id
}

