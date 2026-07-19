#!/bin/zsh
# Aethero Blog Image Generator — macOS launcher (dvojklik v Finderu)

cd "$(dirname "$0")"

# Nainstaluj requests pokud chybí
python3 -c "import requests" 2>/dev/null || pip3 install requests -q

# Zeptej se na název článku přes dialog
TITLE=$(osascript -e 'display dialog "Zadej název blogu článku:" default answer "" with title "Aethero Image Generator" buttons {"Zrušit", "Generovat"} default button "Generovat"' -e 'text returned of result' 2>/dev/null)

if [ -z "$TITLE" ]; then
  echo "Zrušeno."
  exit 0
fi

echo "\nGeneruji obrázek pro: $TITLE\n"
python3 aethero_blog_image.py "$TITLE"

echo "\nStiskni Enter pro zavření..."
read
