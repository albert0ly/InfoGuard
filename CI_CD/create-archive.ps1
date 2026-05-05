# Remove old files
Remove-Item app.zip -ErrorAction SilentlyContinue
Remove-Item deploy-temp -Recurse -ErrorAction SilentlyContinue

# Create temp directory
New-Item -ItemType Directory -Path "deploy-temp" -Force

# Copy only what you need
Copy-Item "dist" -Destination "deploy-temp\dist" -Recurse
Copy-Item "assets" -Destination "deploy-temp\assets" -Recurse
Copy-Item "package.json" -Destination "deploy-temp\"
Copy-Item "package-lock.json" -Destination "deploy-temp\"
Copy-Item ".deployment" -Destination "deploy-temp\"

# Create zip
Compress-Archive -Path "deploy-temp\*" -DestinationPath "app.zip" -Force

# Clean up
Remove-Item deploy-temp -Recurse -Force

Write-Host "✅ Created app.zip with only deployment files"