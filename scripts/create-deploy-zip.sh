#!/bin/bash

PROJECT_DIR="/c/Users/albertly/source/repos/albert0ly/InfoGuard01"
cd "$PROJECT_DIR" || exit 1

echo "Current directory: $(pwd)"

# Clean up
rm -rf deploy-temp app.zip

# Create temp directory
mkdir -p deploy-temp

# Copy files
echo "Copying files..."
cp -r dist deploy-temp/ 2>/dev/null || echo "No dist folder"
cp -r assets deploy-temp/ 2>/dev/null || echo "No assets folder"
cp package.json deploy-temp/ 2>/dev/null || echo "No package.json"
cp package-lock.json deploy-temp/ 2>/dev/null || echo "No package-lock.json"

# Show what we have
echo ""
echo "=== Files in deploy-temp ==="
ls -lR deploy-temp/

# Create zip - try PowerShell method from Git Bash
echo ""
echo "Creating zip using PowerShell..."
powershell.exe -Command "Compress-Archive -Path deploy-temp\\* -DestinationPath app.zip -Force"

# Verify
echo ""
if [ -f "app.zip" ]; then
    echo "=== Success! ==="
    ls -lh app.zip
else
    echo "=== FAILED - app.zip not created ==="
fi

# Clean up
rm -rf deploy-temp