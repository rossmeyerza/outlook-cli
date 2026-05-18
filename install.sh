#!/usr/bin/env bash
# outlook-cli installer and updater
# Usage: curl -fsSL https://outlook-cli.21436587.xyz/install.sh | bash
# Or run directly after download.

set -euo pipefail

BLUE='\033[0;34m'
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
BOLD='\033[1m'
NC='\033[0m'

info()  { printf "${BLUE}==>$NC %s\n" "$1"; }
ok()    { printf "${GREEN}ok$NC %s\n" "$1"; }
warn()  { printf "${YELLOW}warn$NC %s\n" "$1"; }
die()   { printf "${RED}error$NC %s\n" "$1"; exit 1; }
step()  { printf "\n${BOLD}%s${NC}\n" "$1"; }

INSTALL_DIR="${OUTLOOK_CLI_DIR:-${HOME}/.local/lib/outlook-draft-cli}"
BIN_DIR="${HOME}/.local/bin"
IS_UPGRADE=false

# ── Detect upgrade vs fresh install ───────────────────────────────

if [ -d "${INSTALL_DIR}/.git" ]; then
  IS_UPGRADE=true
fi

# ── Header ────────────────────────────────────────────────────────

printf "\n${BOLD}outlook-cli installer${NC}\n"
if [ "$IS_UPGRADE" = true ]; then
  printf "Upgrading existing install at %s\n\n" "$INSTALL_DIR"
else
  printf "Fresh install to %s\n\n" "$INSTALL_DIR"
fi

# ── Prerequisites ──────────────────────────────────────────────────

step "Checking prerequisites"

if ! command -v python3 >/dev/null 2>&1; then
  die "python3 is required. Install Python 3.12+ and try again."
fi

PY_VERSION=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
PY_MAJOR=$(echo "$PY_VERSION" | cut -d. -f1)
PY_MINOR=$(echo "$PY_VERSION" | cut -d. -f2)
if [ "$PY_MAJOR" -lt 3 ] || { [ "$PY_MAJOR" -eq 3 ] && [ "$PY_MINOR" -lt 12 ]; }; then
  die "Python 3.12+ is required. Found Python ${PY_VERSION}."
fi
ok "Python ${PY_VERSION}"

if ! command -v git >/dev/null 2>&1; then
  die "git is required. Install git and try again."
fi
ok "git $(git --version | awk '{print $3}')"

mkdir -p "$BIN_DIR"
mkdir -p "$(dirname "$INSTALL_DIR")"

# ── Clone or pull ─────────────────────────────────────────────────

step "Fetching outlook-cli"

if [ "$IS_UPGRADE" = true ]; then
  info "Pulling latest changes..."
  BEFORE=$(git -C "$INSTALL_DIR" rev-parse --short HEAD)
  git -C "$INSTALL_DIR" pull --ff-only
  AFTER=$(git -C "$INSTALL_DIR" rev-parse --short HEAD)
  if [ "$BEFORE" = "$AFTER" ]; then
    ok "Already up to date ($AFTER)"
  else
    ok "Updated $BEFORE -> $AFTER"
  fi
else
  info "Cloning repository..."
  git clone https://github.com/rossmeyerza/outlook-draft-cli.git "$INSTALL_DIR"
  ok "Cloned"
fi

# ── Virtual environment and dependencies ──────────────────────────

step "Installing dependencies"

python3 -m venv "${INSTALL_DIR}/.venv"
"${INSTALL_DIR}/.venv/bin/pip" install --quiet --upgrade pip
"${INSTALL_DIR}/.venv/bin/pip" install --quiet -e "$INSTALL_DIR"
ok "Python packages installed"

"${INSTALL_DIR}/.venv/bin/python" -m playwright install chromium --quiet 2>/dev/null || \
  "${INSTALL_DIR}/.venv/bin/python" -m playwright install chromium
ok "Playwright Chromium ready"

# ── Symlink ────────────────────────────────────────────────────────

ln -sf "${INSTALL_DIR}/.venv/bin/outlook-cli" "${BIN_DIR}/outlook-cli"
ok "outlook-cli linked to ${BIN_DIR}/outlook-cli"

# ── Configuration (fresh install only) ────────────────────────────

if [ "$IS_UPGRADE" = false ]; then
  step "Configuration"

  printf "\noutlook-cli needs your Microsoft 365 email and password\n"
  printf "to authenticate via your organisation's SSO.\n"
  printf "These are stored only in %s/.env\n\n" "$INSTALL_DIR"

  # Read from /dev/tty so this works even when piped via curl | bash
  exec 3</dev/tty

  while true; do
    printf "MS_EMAIL (your work email): "
    read -r MS_EMAIL <&3
    if echo "$MS_EMAIL" | grep -qE '^[^@]+@[^@]+\.[^@]+$'; then
      break
    fi
    printf "Please enter a valid email address.\n"
  done

  while true; do
    printf "MS_PASSWORD (input hidden): "
    stty -echo <&3
    read -r MS_PASSWORD <&3
    stty echo <&3
    printf "\n"
    if [ -n "$MS_PASSWORD" ]; then
      break
    fi
    printf "Password cannot be empty.\n"
  done

  printf "LOCAL_TIMEZONE (default: Europe/London): "
  read -r LOCAL_TIMEZONE <&3
  LOCAL_TIMEZONE="${LOCAL_TIMEZONE:-Europe/London}"

  printf "OUTLOOK_TIMEZONE (default: GMT Standard Time): "
  read -r OUTLOOK_TIMEZONE <&3
  OUTLOOK_TIMEZONE="${OUTLOOK_TIMEZONE:-GMT Standard Time}"

  exec 3<&-

  cat > "${INSTALL_DIR}/.env" << EOF
MS_EMAIL=${MS_EMAIL}
MS_PASSWORD=${MS_PASSWORD}
LOCAL_TIMEZONE=${LOCAL_TIMEZONE}
OUTLOOK_TIMEZONE=${OUTLOOK_TIMEZONE}
SIGNATURE_NEW_FILE=signature-new.html
SIGNATURE_REPLY_FILE=signature-reply.html
EOF

  ok ".env written to ${INSTALL_DIR}/.env"
else
  ok ".env already exists, not modified"
fi

# ── PATH check ────────────────────────────────────────────────────

if ! echo ":${PATH}:" | grep -q ":${BIN_DIR}:"; then
  warn "${BIN_DIR} is not in your PATH."
  warn "Add this to your shell profile (~/.bashrc, ~/.zshrc, etc.):"
  warn "  export PATH=\"${BIN_DIR}:\$PATH\""
fi

# ── Done ──────────────────────────────────────────────────────────

printf "\n"
if [ "$IS_UPGRADE" = true ]; then
  printf "${GREEN}${BOLD}outlook-cli updated successfully.${NC}\n\n"
  printf "  Run: outlook-cli --help\n"
else
  printf "${GREEN}${BOLD}outlook-cli installed successfully.${NC}\n\n"
  printf "  Next: run outlook-cli auth to sign in\n"
  printf "  Then: outlook-cli config check\n"
fi
printf "\n"
