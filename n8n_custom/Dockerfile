FROM node:20-bookworm-slim

USER root

RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    python3-venv \
    bash \
    coreutils \
    ca-certificates \
    dumb-init \
    && python3 -m venv /opt/venv \
    && /opt/venv/bin/pip install --no-cache-dir openpyxl python-docx pandas \
    && npm install -g n8n \
    && rm -rf /var/lib/apt/lists/*

ENV PATH="/opt/venv/bin:$PATH"
ENV N8N_USER_FOLDER="/home/node/.n8n"
ENV NODE_ENV="production"

RUN mkdir -p /home/node/.n8n /data \
    && chown -R node:node /home/node /data /opt/venv

USER node

ENTRYPOINT ["dumb-init", "--"]
CMD ["n8n", "start"]