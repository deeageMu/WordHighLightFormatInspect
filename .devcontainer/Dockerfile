FROM mcr.microsoft.com/devcontainers/base:bullseye

USER root

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        python3 \
        python3-pip \
        python3-venv \
        docker.io \
        git-lfs \
    && git lfs install --system \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /tmp/requirements.txt

RUN python3 -m pip install --no-cache-dir --upgrade pip \
    && python3 -m pip install --no-cache-dir -r /tmp/requirements.txt

USER vscode