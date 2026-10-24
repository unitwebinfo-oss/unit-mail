#!/bin/sh
# Postinst completo — o "afterInstall" do electron-builder SUBSTITUI o postinst
# gerado automaticamente, então este script precisa refazer o que o padrão faz:
# lançador em /usr/bin, caches de desktop/mime — além das correções do app.

set -e

# Bug conhecido do Electron/Chromium: o SUID sandbox não inicia quando o
# caminho de instalação contém espaço ("/opt/Unit Mail") — execvp corta o
# path no espaço e o zygote morre com FATAL. O app roda com --no-sandbox.
APP="/opt/Unit Mail/unit-mail"

# Lançador de linha de comando (wrapper com --no-sandbox)
if [ -f "$APP" ]; then
  cat > /usr/bin/unit-mail <<'EOF'
#!/bin/sh
exec "/opt/Unit Mail/unit-mail" --no-sandbox "$@"
EOF
  chmod 755 /usr/bin/unit-mail
fi

# O helper do sandbox exige SUID (4755); mantém correto para quando for usado.
chmod 4755 "/opt/Unit Mail/chrome-sandbox" 2>/dev/null || true

# Entrada de menu: mesma correção --no-sandbox na linha Exec
if [ -f /usr/share/applications/unit-mail.desktop ]; then
  sed -i 's|^Exec=.*|Exec="/opt/Unit Mail/unit-mail" --no-sandbox %U|' \
    /usr/share/applications/unit-mail.desktop 2>/dev/null || true
fi

update-mime-database /usr/share/mime >/dev/null 2>&1 || true
update-desktop-database /usr/share/applications >/dev/null 2>&1 || true
exit 0
