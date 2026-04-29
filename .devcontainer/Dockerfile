# syntax=docker/dockerfile:1

FROM mcr.microsoft.com/devcontainers/dotnet:1-8.0-bookworm

ARG DEBIAN_FRONTEND=noninteractive

ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1 \
    ASPNETCORE_URLS=http://0.0.0.0:2805

WORKDIR /workspaces/SmartOffice.Hub

# 未來 package 區：
# 額外 package 請集中放在這裡，讓 full devcontainer mode 與未來共用
# build image 可以從同一個文件化 baseline 演進。
#
# 範例：
# RUN apt-get update \
#     && apt-get install -y --no-install-recommends <package-name> \
#     && rm -rf /var/lib/apt/lists/*
