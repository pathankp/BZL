#!/bin/bash

# Check if running as root
if [ "$(id -u)" != "0" ]; then
  if command -v sudo >/dev/null 2>&1; then
    exec sudo "$0" "$@"
  else
    echo "This script must be run as root. Please either:"
    echo "1. Run this script as root (su root)"
    echo "2. Install sudo and run with sudo"
    exit 1
  fi
fi

# Define default values
version=0.0.1
PORT=8090                              # Default port
GITHUB_PROXY_URL="https://ghfast.top/" # Default proxy URL

# Function to ensure the proxy URL ends with a /
ensure_trailing_slash() {
  if [ -n "$1" ]; then
    case "$1" in
    */) echo "$1" ;;
    *) echo "$1/" ;;
    esac
  else
    echo "$1"
  fi
}

# Ensure the proxy URL ends with a /
GITHUB_PROXY_URL=$(ensure_trailing_slash "$GITHUB_PROXY_URL")

# Read command line options
while getopts ":uhp:c:" opt; do
  case $opt in
  u) UNINSTALL="true" ;;
  h)
    printf "ServerSentry Hub installation script\n\n"
    printf "Usage: ./install-hub.sh [options]\n\n"
    printf "Options: \n"
    printf "  -u  : Uninstall the ServerSentry Hub\n"
    printf "  -p <port> : Specify a port number (default: 8090)\n"
    printf "  -c <url>  : Use a custom GitHub mirror URL (e.g., https://ghfast.top/)\n"
    echo "  -h  : Display this help message"
    exit 0
    ;;
  p) PORT=$OPTARG ;;
  c) GITHUB_PROXY_URL=$(ensure_trailing_slash "$OPTARG") ;;
  \?)
    echo "Invalid option: -$OPTARG"
    exit 1
    ;;
  esac
done

if [ "$UNINSTALL" = "true" ]; then
  # Stop and disable the ServerSentry Hub service
  echo "Stopping and disabling the ServerSentry Hub service..."
  systemctl stop serversentry-hub.service
  systemctl disable serversentry-hub.service

  # Remove the systemd service file
  echo "Removing the systemd service file..."
  rm /etc/systemd/system/serversentry-hub.service

  # Reload the systemd daemon
  echo "Reloading the systemd daemon..."
  systemctl daemon-reload

  # Remove the ServerSentry Hub binary and data
  echo "Removing the ServerSentry Hub binary and data..."
  rm -rf /opt/serversentry

  # Remove the dedicated user
  echo "Removing the dedicated user..."
  userdel serversentry

  echo "The ServerSentry Hub has been uninstalled successfully!"
  exit 0
fi

# Function to check if a package is installed
package_installed() {
  command -v "$1" >/dev/null 2>&1
}

# Check for package manager and install necessary packages if not installed
if package_installed apt-get; then
  if ! package_installed tar || ! package_installed curl; then
    apt-get update
    apt-get install -y tar curl
  fi
elif package_installed yum; then
  if ! package_installed tar || ! package_installed curl; then
    yum install -y tar curl
  fi
elif package_installed pacman; then
  if ! package_installed tar || ! package_installed curl; then
    pacman -Sy --noconfirm tar curl
  fi
else
  echo "Warning: Please ensure 'tar' and 'curl' are installed."
fi

# Create a dedicated user for the service if it doesn't exist
if ! id -u serversentry >/dev/null 2>&1; then
  echo "Creating a dedicated user for the ServerSentry Hub service..."
  useradd -M -s /bin/false serversentry
fi

# Download and install the ServerSentry Hub
echo "Downloading and installing the ServerSentry Hub..."
curl -sL "${GITHUB_PROXY_URL}https://github.com/nak-ventures/serversentry/releases/latest/download/serversentry_$(uname -s)_$(uname -m | sed 's/x86_64/amd64/' | sed 's/armv7l/arm/' | sed 's/aarch64/arm64/').tar.gz" | tar -xz -O serversentry | tee ./serversentry >/dev/null && chmod +x serversentry
mkdir -p /opt/serversentry/serversentry_data
mv ./serversentry /opt/serversentry/serversentry
chown -R serversentry:serversentry /opt/serversentry

# Create the systemd service
printf "Creating the systemd service for the ServerSentry Hub...\n\n"
tee /etc/systemd/system/serversentry-hub.service <<EOF
[Unit]
Description=ServerSentry Hub Service
After=network.target

[Service]
ExecStart=/opt/serversentry/serversentry serve --http "0.0.0.0:$PORT"
WorkingDirectory=/opt/serversentry
User=serversentry
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

# Load and start the service
printf "\nLoading and starting the ServerSentry Hub service...\n"
systemctl daemon-reload
systemctl enable serversentry-hub.service
systemctl start serversentry-hub.service

# Wait for the service to start or fail
sleep 2

# Check if the service is running
if [ "$(systemctl is-active serversentry-hub.service)" != "active" ]; then
  echo "Error: The ServerSentry Hub service is not running."
  echo "$(systemctl status serversentry-hub.service)"
  exit 1
fi

echo "The ServerSentry Hub has been installed and configured successfully! It is now accessible on port $PORT."
