#!/bin/sh
# O helper do sandbox do Chromium exige SUID (4755). Em alguns builds o pacote
# instala como 0755 e o Electron aborta na abertura — garante o bit aqui.
chmod 4755 "/opt/Unit Mail/chrome-sandbox" 2>/dev/null || true
exit 0
