#!/bin/bash
# ==========================================================================
#   Script de démarrage du Dashboard RH
# ==========================================================================

echo ""
echo "=========================================="
echo "🏢 DASHBOARD RH 2026"
echo "=========================================="
echo ""

# Aller dans le répertoire du script
cd "$(dirname "$0")"

# Vérifier Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 n'est pas installé!"
    echo "   Installez-le avec: sudo apt install python3 python3-pip"
    exit 1
fi

# Vérifier les dépendances
echo "📦 Vérification des dépendances..."
python3 -c "import flask" 2>/dev/null || {
    echo "⏳ Installation de Flask..."
    pip3 install flask --quiet
}

python3 -c "import pandas" 2>/dev/null || {
    echo "⏳ Installation de Pandas..."
    pip3 install pandas openpyxl --quiet
}

# Vérifier les certificats SSL
if [ -f "cert.pem" ] && [ -f "key.pem" ]; then
    echo "✅ Certificats SSL trouvés"
else
    echo "⚠️  Certificats SSL non trouvés"
    echo "   Pour générer un certificat auto-signé:"
    echo "   python3 generate_cert.py"
    echo ""
fi

# Vérifier les données
if [ -f "data/data.json" ]; then
    echo "✅ Fichier de données trouvé"
else
    echo "⚠️  Fichier data.json non trouvé"
    if [ -f "data/TDB_COURANT.xlsx" ]; then
        echo "⏳ Génération des données depuis Excel..."
        python3 -c "from app import generate_data_from_excel; print(generate_data_from_excel())"
    else
        echo "❌ Aucun fichier Excel trouvé dans data/"
        echo "   Déposez TDB_COURANT.xlsx dans le dossier data/"
    fi
fi

echo ""
echo "🚀 Démarrage du serveur..."
echo ""

# Démarrer l'application
python3 app.py
