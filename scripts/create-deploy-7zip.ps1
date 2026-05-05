cd C:\Users\albertly\source\repos\albert0ly\InfoGuard01

# Clean up
Remove-Item app.zip -ErrorAction SilentlyContinue
Remove-Item deploy-temp -Recurse -ErrorAction SilentlyContinue

# Create temp dir with Linux-style structure
New-Item -ItemType Directory -Path "deploy-temp" -Force | Out-Null
Copy-Item "dist" -Destination "deploy-temp\" -Recurse
Copy-Item "assets" -Destination "deploy-temp\" -Recurse  
Copy-Item "package.json" -Destination "deploy-temp\"
Copy-Item "package-lock.json" -Destination "deploy-temp\"

# Use 7z in "store paths" mode
& "C:\Program Files\7-Zip\7z.exe" a -tzip app.zip .\deploy-temp\* -r

# Clean up
Remove-Item deploy-temp -Recurse -Force