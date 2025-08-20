#!/bin/bash
set -e

echo "Installing ServerSentry Hub from source..."

# --- Install dependencies ---
echo "Updating packages and installing dependencies..."
sudo apt-get update -y
sudo apt-get install -y curl git build-essential nodejs npm golang

# --- Prepare install directory ---
sudo rm -rf /opt/serversentry
sudo mkdir -p /opt/serversentry
sudo chown $(whoami) /opt/serversentry

# --- Clone repo ---
echo "Cloning ServerSentry..."
git clone https://github.com/pathankp/BZL.git /opt/serversentry/src

# --- Build frontend ---
echo "Building UI..."
cd /opt/serversentry/src/ui
npm install
npm run build

# --- Build backend ---
echo "Building Go backend..."
cd /opt/serversentry/src
go build -o serversentry ./beszel

# --- Move binary to /usr/local/bin ---
sudo mv /opt/serversentry/src/serversentry /usr/local/bin/serversentry
sudo chmod +x /usr/local/bin/serversentry

# --- Create service user if not exists ---
if ! id "serversentry" &>/dev/null; then
    sudo useradd -r -s /bin/false serversentry
fi

# --- Create systemd service ---
echo "Creating systemd service..."
sudo tee /etc/systemd/system/serversentry-hub.service > /dev/null <<EOL
[Unit]
Description=ServerSentry Hub Service
After=network.target

[Service]
ExecStart=/usr/local/bin/serversentry serve --http "0.0.0.0:8090"
WorkingDirectory=/opt/serversentry
User=serversentry
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOL

# --- Reload systemd and start ---
echo "Starting ServerSentry Hub service..."
sudo systemctl daemon-reload
sudo systemctl enable serversentry-hub
sudo systemctl start serversentry-hub

echo "Installation complete! ServerSentry Hub is running on port 8090."
