resource "azurerm_resource_group" "main" {
  name     = "rg-${var.project_name}-${var.environment}"
  location = var.location
  tags     = merge(var.tags, { Environment = var.environment })
}

resource "azurerm_service_plan" "main" {
  name                = "asp-${var.project_name}-${var.environment}"
  resource_group_name = azurerm_resource_group.main.name
  location            = azurerm_resource_group.main.location
  os_type             = "Linux"
  sku_name            = var.app_service_sku
  tags                = merge(var.tags, { Environment = var.environment })
}

resource "azurerm_linux_web_app" "backend" {
  name                = "app-${var.project_name}-backend-${var.environment}"
  resource_group_name = azurerm_resource_group.main.name
  location            = azurerm_resource_group.main.location
  service_plan_id     = azurerm_service_plan.main.id
  https_only          = true
  tags                = merge(var.tags, { Environment = var.environment })

  site_config {
    always_on = var.environment == "prod" ? true : false
    
    application_stack {
      node_version = var.node_version
    }

    cors {
      allowed_origins     = var.allowed_origins
      support_credentials = true
    }

    ftps_state = "FtpsOnly"
    health_check_path = "/health"
  }

  app_settings = {
    "WEBSITE_NODE_DEFAULT_VERSION" = "~18"
    "NODE_ENV"                     = var.environment
    "APPLICATIONINSIGHTS_CONNECTION_STRING" = azurerm_application_insights.main.connection_string
  }

  identity {
    type = "SystemAssigned"
  }
}

resource "azurerm_log_analytics_workspace" "main" {
  name                = "log-${var.project_name}-${var.environment}"
  resource_group_name = azurerm_resource_group.main.name
  location            = azurerm_resource_group.main.location
  sku                 = "PerGB2018"
  retention_in_days   = 30
  tags                = merge(var.tags, { Environment = var.environment })
}

resource "azurerm_application_insights" "main" {
  name                = "appi-${var.project_name}-${var.environment}"
  resource_group_name = azurerm_resource_group.main.name
  location            = azurerm_resource_group.main.location
  workspace_id        = azurerm_log_analytics_workspace.main.id
  application_type    = "Node.JS"
  tags                = merge(var.tags, { Environment = var.environment })
}

resource "azurerm_storage_account" "addin" {
  name                     = lower("st${replace(var.project_name, "-", "")}${var.environment}")
  resource_group_name      = azurerm_resource_group.main.name
  location                 = azurerm_resource_group.main.location
  account_tier             = "Standard"
  account_replication_type = "LRS"
  
  blob_properties {
    cors_rule {
      allowed_origins    = ["https://outlook.office.com", "https://outlook.office365.com"]
      allowed_methods    = ["GET", "HEAD"]
      allowed_headers    = ["*"]
      exposed_headers    = ["*"]
      max_age_in_seconds = 3600
    }
  }
  
  tags = merge(var.tags, { Environment = var.environment })
}

resource "azurerm_storage_container" "manifests" {
  name                  = "manifests"
  storage_account_name  = azurerm_storage_account.addin.name
  container_access_type = "blob"
}

data "azurerm_client_config" "current" {}

resource "azurerm_key_vault" "main" {
  name                       = "kv-${var.project_name}-${substr(var.environment, 0, 3)}-${substr(sha256(data.azurerm_client_config.current.subscription_id), 0, 4)}"
  resource_group_name        = azurerm_resource_group.main.name
  location                   = azurerm_resource_group.main.location
  tenant_id                  = data.azurerm_client_config.current.tenant_id
  sku_name                   = "standard"
  soft_delete_retention_days = 7
  purge_protection_enabled   = false
  tags                       = merge(var.tags, { Environment = var.environment })
}

resource "azurerm_key_vault_access_policy" "app_service" {
  key_vault_id = azurerm_key_vault.main.id
  tenant_id    = data.azurerm_client_config.current.tenant_id
  object_id    = azurerm_linux_web_app.backend.identity[0].principal_id

  secret_permissions = [
    "Get",
    "List"
  ]
}

resource "azurerm_key_vault_access_policy" "current_user" {
  key_vault_id = azurerm_key_vault.main.id
  tenant_id    = data.azurerm_client_config.current.tenant_id
  object_id    = data.azurerm_client_config.current.object_id

  secret_permissions = [
    "Get",
    "List",
    "Set",
    "Delete",
    "Recover",
    "Backup",
    "Restore"
  ]
}