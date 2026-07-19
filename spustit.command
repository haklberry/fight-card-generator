#!/bin/bash
# Dvojklikem spustíš aplikaci — otevře se v prohlížeči automaticky

cd "$(dirname "$0")"

echo "=== Fight Card Generator ==="
echo ""

# Najdi Python s nainstalovaným streamlit, nebo použij výchozí python3
PYTHON=""
for candidate in \
    "/Library/Developer/CommandLineTools/Library/Frameworks/Python3.framework/Versions/3.9/bin/python3.9" \
    "$(which python3 2>/dev/null)" \
    "$(which python 2>/dev/null)"; do
    if [ -n "$candidate" ] && "$candidate" -c "import streamlit" 2>/dev/null; then
        PYTHON="$candidate"
        break
    fi
done

if [ -z "$PYTHON" ]; then
    # Žádný Python nemá streamlit — zkus ho nainstalovat
    PY=$(which python3 2>/dev/null || which python 2>/dev/null)
    if [ -z "$PY" ]; then
        echo "❌ Python 3 není nainstalován."
        echo "   Stáhni ho z https://www.python.org a nainstaluj."
        read -p "Stiskni Enter pro zavření..."
        exit 1
    fi
    PYTHON="$PY"
fi

echo "✅ Python: $($PYTHON --version)"
echo ""
echo "📦 Instaluji závislosti (jen při prvním spuštění)..."
"$PYTHON" -m pip install -r requirements.txt -q 2>/dev/null \
    || "$PYTHON" -m pip install -r requirements.txt -q --break-system-packages

echo ""
echo "🚀 Spouštím aplikaci (otevře se prohlížeč)..."
echo "   Pro ukončení zavři toto okno nebo stiskni Ctrl+C"
echo ""

"$PYTHON" -m streamlit run app_karta.py \
    --server.headless false \
    --browser.gatherUsageStats false
