#!/bin/bash
# update_server.sh
# -----------------
# Run this on the server (via the DigitalOcean Web Console) AFTER pulling
# the latest code with `git pull`. It:
#   1. Points nginx at the new public landing page (public/index.html) for
#      the main domain, and reverse-proxies /app/ to the Streamlit tool.
#   2. Tells Streamlit to run under the /app/ base path so it works
#      correctly behind that proxy.
#   3. Restarts both services.
#
# Usage:
#   cd /opt/reco-tool && git pull && bash update_server.sh
set -e

APP_DIR="/opt/reco-tool"
cd "$APP_DIR"

# --- nginx: serve the public landing page at "/", proxy "/app/" to Streamlit ---
cat > /etc/nginx/sites-available/reco-tool << 'NGINX'
server {
    listen 80;
    server_name recomatrix.com www.recomatrix.com _;

    client_max_body_size 200M;

    root /opt/reco-tool/public;
    index index.html;

    location / {
        try_files $uri $uri/ =404;
    }

    location /app/ {
        proxy_pass http://127.0.0.1:8501/app/;
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

nginx -t

# --- systemd: run Streamlit under the /app base path ---
cat > /etc/systemd/system/reco-tool.service << 'UNIT'
[Unit]
Description=Reco Tool Streamlit App
After=network.target

[Service]
Type=simple
WorkingDirectory=/opt/reco-tool
ExecStart=/opt/reco-tool/venv/bin/streamlit run app.py --server.port=8501 --server.address=127.0.0.1 --server.headless=true --server.baseUrlPath=app
Restart=always
RestartSec=5
User=root

[Install]
WantedBy=multi-user.target
UNIT

systemctl daemon-reload
systemctl restart reco-tool
systemctl reload nginx

echo "=== UPDATE COMPLETE ==="
echo "Public landing page: http://recomatrix.com/"
echo "Login / portal:      http://recomatrix.com/app/"
