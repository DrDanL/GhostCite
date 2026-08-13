#!/usr/bin/env bash
set -euo pipefail

echo "🔨 Building and starting GhostCite..."
docker compose up --build -d

echo ""
echo "✅ GhostCite is running!"
echo ""
echo "   Open the app in your browser:"
echo ""
echo "   👉 http://localhost:8000"
echo ""
echo "   To view logs:   docker compose logs -f"
echo "   To stop:        docker compose down"
