#!/bin/bash

PORT=45876
KEY=""
TOKEN=""
HUB_URL=""

usage() {
  printf "ServerSentry Agent homebrew installation script\n\n"
  printf "Usage: ./install-agent-brew.sh [options]\n\n"
  printf "Options: \n"
  printf "  -k            SSH key (required, or interactive if not provided)\n"
  printf "  -p            Port (default: $PORT)\n"
  printf "  -t            Token (optional for backwards compatibility)\n"
  printf "  -url          Hub URL (optional for backwards compatibility)\n"
  printf "  -h, --help    Display this help message\n"
  exit 0
}

# Parse arguments
while [ $# -gt 0 ]; do
  case "$1" in
  -k)
    shift
    KEY="$1"
    ;;
  -p)
    shift
    PORT="$1"
    ;;
  -t)
    shift
    TOKEN="$1"
    ;;
  -url)
    shift
    HUB_URL="$1"
    ;;
  -h | --help)
    usage
    ;;
  *)
    echo "Invalid option: $1" >&2
    usage
    ;;
  esac
  shift
done

# Check if brew is installed, prompt to install if not
if ! command -v brew &>/dev/null; then
  read -p "Homebrew is not installed. Would you like to install it now? (y/n): " install_brew
  if [[ $install_brew =~ ^[Yy]$ ]]; then
    echo "Installing Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

    # Verify installation was successful
    if ! command -v brew &>/dev/null; then
      echo "Homebrew installation failed. Please install manually and try again."
      exit 1
    fi
    echo "Homebrew installed successfully."
  else
    echo "Homebrew is required. Please install Homebrew and try again."
    exit 1
  fi
fi

if [ -z "$KEY" ]; then
  read -p "Enter SSH key: " KEY
fi

# TOKEN and HUB_URL are optional for backwards compatibility - no interactive prompts

mkdir -p ~/.config/serversentry ~/.cache/serversentry

echo "KEY=\"$KEY\"" >~/.config/serversentry/serversentry-agent.env
echo "LISTEN=$PORT" >>~/.config/serversentry/serversentry-agent.env

if [ -n "$TOKEN" ]; then
  echo "TOKEN=\"$TOKEN\"" >>~/.config/serversentry/serversentry-agent.env
fi
if [ -n "$HUB_URL" ]; then
  echo "HUB_URL=\"$HUB_URL\"" >>~/.config/serversentry/serversentry-agent.env
fi

brew tap nak-ventures/serversentry
brew install serversentry-agent
brew services start serversentry-agent

printf "\nCheck status: brew services info serversentry-agent\n"
echo "Stop: brew services stop serversentry-agent"
echo "Start: brew services start serversentry-agent"
echo "Restart: brew services restart serversentry-agent"
echo "Upgrade: brew upgrade serversentry-agent"
echo "Uninstall: brew uninstall serversentry-agent"
echo "View logs in ~/.cache/serversentry/serversentry-agent.log"
printf "Change environment variables in ~/.config/serversentry/serversentry-agent.env\n"
