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
FORCE_RECONFIGURE=false

for arg in "$@"; do
  case "$arg" in
    --reconfigure|--reconfig)
      FORCE_RECONFIGURE=true
      ;;
  esac
done

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

if ! python3 -c 'import ensurepip' >/dev/null 2>&1; then
  printf "${RED}error${NC} Python venv module is missing.\n"
  printf "On Debian/Ubuntu/WSL, install it with:\n"
  printf "  sudo apt install python%s-venv python3-venv\n" "${PY_VERSION}"
  printf "Then re-run this installer.\n"
  exit 1
fi

rm -rf "${INSTALL_DIR}/.venv"
python3 -m venv "${INSTALL_DIR}/.venv"
"${INSTALL_DIR}/.venv/bin/pip" install --quiet --upgrade pip
"${INSTALL_DIR}/.venv/bin/pip" install --quiet -e "$INSTALL_DIR"
ok "Python packages installed"

"${INSTALL_DIR}/.venv/bin/python" -m playwright install chromium --quiet 2>/dev/null || \
  "${INSTALL_DIR}/.venv/bin/python" -m playwright install chromium
ok "Playwright Chromium ready"

# ── WSL note for headless deps ────────────────────────────────────

if grep -qi microsoft /proc/version 2>/dev/null; then
  warn "Detected WSL. If headless auth fails, run:"
  warn "  sudo ${INSTALL_DIR}/.venv/bin/python -m playwright install-deps chromium"
fi

# ── Symlink ────────────────────────────────────────────────────────

ln -sf "${INSTALL_DIR}/.venv/bin/outlook-cli" "${BIN_DIR}/outlook-cli"
ok "outlook-cli linked to ${BIN_DIR}/outlook-cli"

# ── Configuration (fresh install only) ────────────────────────────

# ── Configuration ─────────────────────────────────────────────────

env_value() {
  # Print the raw value of a key from .env, empty if missing
  local key="$1"
  if [ -f "${INSTALL_DIR}/.env" ]; then
    awk -F= -v k="$key" '$1 == k { sub(/^[^=]+=/, ""); print; exit }' "${INSTALL_DIR}/.env"
  fi
}

MS_EMAIL_CURRENT="$(env_value MS_EMAIL || true)"
MS_PASSWORD_CURRENT="$(env_value MS_PASSWORD || true)"
MS_EMAIL_PLACEHOLDER="your.email@company.com"
MS_PASSWORD_PLACEHOLDER="your-password"

NEEDS_CONFIG=false
if [ "$FORCE_RECONFIGURE" = true ]; then
  NEEDS_CONFIG=true
fi
if [ ! -f "${INSTALL_DIR}/.env" ]; then
  NEEDS_CONFIG=true
fi
if [ -z "$MS_EMAIL_CURRENT" ] || [ "$MS_EMAIL_CURRENT" = "$MS_EMAIL_PLACEHOLDER" ]; then
  NEEDS_CONFIG=true
fi
if [ -z "$MS_PASSWORD_CURRENT" ] || [ "$MS_PASSWORD_CURRENT" = "$MS_PASSWORD_PLACEHOLDER" ]; then
  NEEDS_CONFIG=true
fi

if [ "$NEEDS_CONFIG" = true ]; then
  step "Configuration"

  if [ "$FORCE_RECONFIGURE" = true ]; then
    info "--reconfigure was passed, prompting for new credentials."
  fi

  printf "\noutlook-cli needs your Microsoft 365 email and password\n"
  printf "to authenticate via your organisation's SSO.\n"
  printf "These are stored only in %s/.env\n\n" "$INSTALL_DIR"

  # Read from /dev/tty so this works even when piped via curl | bash
  exec 3</dev/tty

  while true; do
    if [ -n "$MS_EMAIL_CURRENT" ] && [ "$MS_EMAIL_CURRENT" != "$MS_EMAIL_PLACEHOLDER" ]; then
      printf "MS_EMAIL [%s]: " "$MS_EMAIL_CURRENT"
      read -r MS_EMAIL_INPUT <&3
      MS_EMAIL="${MS_EMAIL_INPUT:-$MS_EMAIL_CURRENT}"
    else
      printf "MS_EMAIL (your work email): "
      read -r MS_EMAIL <&3
    fi
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

  CURRENT_LOCAL_TZ="$(env_value LOCAL_TIMEZONE || true)"
  printf "LOCAL_TIMEZONE (default: %s): " "${CURRENT_LOCAL_TZ:-Europe/London}"
  read -r LOCAL_TIMEZONE <&3
  LOCAL_TIMEZONE="${LOCAL_TIMEZONE:-${CURRENT_LOCAL_TZ:-Europe/London}}"

  CURRENT_OUTLOOK_TZ="$(env_value OUTLOOK_TIMEZONE || true)"
  printf "OUTLOOK_TIMEZONE (default: %s): " "${CURRENT_OUTLOOK_TZ:-GMT Standard Time}"
  read -r OUTLOOK_TIMEZONE <&3
  OUTLOOK_TIMEZONE="${OUTLOOK_TIMEZONE:-${CURRENT_OUTLOOK_TZ:-GMT Standard Time}}"

  exec 3<&-

  cat > "${INSTALL_DIR}/.env" << EOF
MS_EMAIL=${MS_EMAIL}
MS_PASSWORD=${MS_PASSWORD}
LOCAL_TIMEZONE=${LOCAL_TIMEZONE}
OUTLOOK_TIMEZONE=${OUTLOOK_TIMEZONE}
SIGNATURE_NEW_FILE=signature-new.html
SIGNATURE_REPLY_FILE=signature-reply.html
EOF

  chmod 600 "${INSTALL_DIR}/.env"
  ok ".env written to ${INSTALL_DIR}/.env"
else
  ok ".env already configured, not modified (use --reconfigure to rewrite)"
fi

# ── PATH check ────────────────────────────────────────────────────

if ! echo ":${PATH}:" | grep -q ":${BIN_DIR}:"; then
  warn "${BIN_DIR} is not in your PATH."
  warn "Add this to your shell profile (~/.bashrc, ~/.zshrc, etc.):"
  warn "  export PATH=\"${BIN_DIR}:\$PATH\""
  export PATH="${BIN_DIR}:${PATH}"
fi

# ── Auth and signature fetch (fresh install only) ─────────────────

if [ "$IS_UPGRADE" = false ] && [ "$NEEDS_CONFIG" = true ]; then
  step "Authentication"
  printf "\nOpening a browser window to sign in to Outlook.\n"
  printf "Complete the MFA prompt, then your signatures will be fetched automatically.\n\n"
  if "${INSTALL_DIR}/.venv/bin/outlook-cli" auth; then
    ok "Authenticated"
    info "Fetching email signatures..."
    if "${INSTALL_DIR}/.venv/bin/outlook-cli" signature fetch; then
      ok "Signatures saved"
    else
      warn "Signature fetch failed. Run: outlook-cli signature fetch"
    fi
  else
    warn "Authentication failed. Run: outlook-cli auth"
  fi
fi

# ── Done ──────────────────────────────────────────────────────────

printf "\n"
if [ "$IS_UPGRADE" = true ]; then
  printf "${GREEN}${BOLD}outlook-cli updated successfully.${NC}\n\n"
  printf "  Run: outlook-cli --help\n"
else
  printf "${GREEN}${BOLD}outlook-cli installed successfully.${NC}\n\n"
  if [ "$NEEDS_CONFIG" = false ]; then
    printf "  Run: outlook-cli auth to sign in\n"
    printf "  Then: outlook-cli signature fetch\n"
  fi
fi
printf "\n"
