# syntax=docker/dockerfile:1

FROM mcr.microsoft.com/devcontainers/dotnet:1-8.0-bookworm

ARG DEBIAN_FRONTEND=noninteractive
ARG NODE_MAJOR=22

ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1 \
    ASPNETCORE_URLS=http://0.0.0.0:2805

WORKDIR /workspaces/SmartOffice.Hub

# 前端 build 使用 Vite。Vite 7 需要 Node.js 20.19+ 或 22.12+，
# 這裡固定使用 Node.js 22 LTS，讓 devcontainer 與 Quick Mode 使用同一套 toolchain。
RUN rm -f /etc/apt/sources.list.d/yarn.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl gnupg \
    && mkdir -p /etc/apt/keyrings \
    && curl -fsSL https://deb.nodesource.com/gpgkey/nodesource-repo.gpg.key \
        | gpg --dearmor -o /etc/apt/keyrings/nodesource.gpg \
    && echo "deb [signed-by=/etc/apt/keyrings/nodesource.gpg] https://deb.nodesource.com/node_${NODE_MAJOR}.x nodistro main" \
        > /etc/apt/sources.list.d/nodesource.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*
