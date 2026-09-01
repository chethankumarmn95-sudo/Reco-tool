#!/bin/bash
set -e
export DEBIAN_FRONTEND=noninteractive NEEDRESTART_MODE=a

apt-get update -y
apt-get install -y python3 python3-venv python3-pip nginx git

APP_DIR="/opt/reco-tool"
cd "$APP_DIR"

python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt

cat > /etc/systemd/system/reco-tool.service << 'UNIT'
[Unit]
Description=Reco Tool Streamlit App
After=network.target

[Service]
Type=simple
WorkingDirectory=/opt/reco-tool
ExecStart=/opt/reco-tool/venv/bin/streamlit run app.py --server.port=8501 --server.address=127.0.0.1 --server.headless=true
Restart=always
RestartSec=5
User=root

[Install]
WantedBy=multi-user.target
UNIT

systemctl daemon-reload
systemctl enable reco-tool
systemctl restart reco-tool

cat > /etc/nginx/sites-available/reco-tool << 'NGINX'
server {
    listen 80;
    server_name recomatrix.com www.recomatrix.com _;

    location / {
        proxy_pass http://127.0.0.1:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_read_timeout 86400;
    }
}
NGINX

rm -f /etc/nginx/sites-enabled/default
ln -sf /etc/nginx/sites-available/reco-tool /etc/nginx/sites-enabled/reco-tool
nginx -t && systemctl restart nginx

ufw allow OpenSSH
ufw allow 'Nginx Full'
ufw --force enable

echo "=== SETUP COMPLETE - visit http://168.144.88.189 ==="
