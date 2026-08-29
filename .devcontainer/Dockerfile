FROM mcr.microsoft.com/devcontainers/python:3.12

USER root

RUN apt-get update \
    && apt-get install -y --no-install-recommends git-lfs \
    && git lfs install --system \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

USER vscode