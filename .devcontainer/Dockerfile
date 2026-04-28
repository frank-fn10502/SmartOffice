# syntax=docker/dockerfile:1

FROM mcr.microsoft.com/devcontainers/dotnet:1-8.0-bookworm

ARG DEBIAN_FRONTEND=noninteractive

ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1 \
    ASPNETCORE_URLS=http://0.0.0.0:2805

WORKDIR /workspaces/SmartOffice.Hub

# Future package area:
# Keep extra packages here so both full devcontainer mode and future shared
# build images can evolve from one documented baseline.
#
# Example:
# RUN apt-get update \
#     && apt-get install -y --no-install-recommends <package-name> \
#     && rm -rf /var/lib/apt/lists/*
