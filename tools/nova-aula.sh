#!/bin/sh
# Gera, para uma aula, a prévia borrada do mapa, o QR code e a tela final 4K do vídeo.
# uso: tools/nova-aula.sh 02 /caminho/mapa-mental.png /caminho/mapa-mental.pdf
set -e
cd "$(dirname "$0")/.."
A="$1"; PNG="$2"; PDF="$3"

SITE=$(cat tools/site-url.txt)
SDK=$(ls -d /Library/Developer/CommandLineTools/SDKs/MacOSX26*.sdk 2>/dev/null | tail -1)
CHROME=$(find ~/.cache/hyperframes/chrome -name chrome-headless-shell -type f | head -1)

[ -x tools/qr ] || swiftc -O ${SDK:+-sdk "$SDK"} tools/qr.swift -o tools/qr

if [ -n "$PDF" ]; then
  SUF=$(openssl rand -hex 3)
  cp "$PDF" "mapas/aula-$A-$SUF.pdf"
  echo "PDF: mapas/aula-$A-$SUF.pdf  (copie este caminho para aulas.js)"
fi
if [ -n "$PNG" ]; then
  ffmpeg -v error -y -i "$PNG" -filter_complex \
    "[0]scale=1400:-2,split[a][b];[a]gblur=sigma=7[bl];[b]crop=1400:250:0:0[top];[bl][top]overlay=0:0" \
    -q:v 3 "previews/aula-$A.jpg"
fi

URL="${SITE}?a=$A&o=qr"
tools/qr "$URL" "qr/aula-$A.png" 1200

SHORT=$(echo "$SITE" | sed 's|https://||; s|/$||')
"$CHROME" --headless --disable-gpu --allow-file-access-from-files --hide-scrollbars \
  --virtual-time-budget=6000 --window-size=1920,1080 --force-device-scale-factor=2 \
  --screenshot="$PWD/endcard/tela-final-aula-$A-4k.png" \
  "file://$PWD/endcard/endcard.html?a=$A&u=$SHORT" 2>/dev/null
ffmpeg -v error -y -loop 1 -framerate 30 -t 20 -i "endcard/tela-final-aula-$A-4k.png" \
  -c:v libx264 -crf 16 -pix_fmt yuv420p -movflags +faststart "endcard/tela-final-aula-$A-4k.mp4"
echo "QR → $URL"
echo "Tela final: endcard/tela-final-aula-$A-4k.png (+ .mp4 de 20 s)"
echo "Link para a descrição do YouTube: ${SITE}?a=$A&o=yt"
