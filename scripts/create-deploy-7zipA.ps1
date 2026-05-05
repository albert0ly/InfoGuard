cd C:\Users\albertly\source\repos\albert0ly\InfoGuard01

npm run build:prod

Remove-Item app.zip -ErrorAction SilentlyContinue
Remove-Item deploy-temp -Recurse -ErrorAction SilentlyContinue

New-Item -ItemType Directory -Path "deploy-temp" -Force | Out-Null

# Copy CONTENTS of dist to root of deploy-temp
Copy-Item "dist\*" -Destination "deploy-temp\" -Recurse

# Copy assets
Copy-Item "assets" -Destination "deploy-temp\assets" -Recurse

# Copy package files
Copy-Item "package.json" -Destination "deploy-temp\"
Copy-Item "package-lock.json" -Destination "deploy-temp\"

# Create zip
& "C:\Program Files\7-Zip\7z.exe" a -tzip app.zip .\deploy-temp\* -r

Remove-Item deploy-temp -Recurse -Force