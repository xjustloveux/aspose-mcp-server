# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Version from build arg (passed from CI/CD, may include 'v' prefix)
ARG VERSION=1.0.0

# Copy project file and restore dependencies
# Pass explicit version to prevent git tag auto-detection
COPY AsposeMcpServer.csproj .
RUN VERSION_NUM=$(echo "${VERSION}" | sed 's/^v//') && \
    dotnet restore AsposeMcpServer.csproj -p:Version=${VERSION_NUM} -p:VersionExplicitlySet=true

# Copy source code and build
COPY . .
RUN VERSION_NUM=$(echo "${VERSION}" | sed 's/^v//') && \
    dotnet publish AsposeMcpServer.csproj -c Release -o /app/publish --no-restore -p:Version=${VERSION_NUM} -p:VersionExplicitlySet=true -p:DebugType=none && \
    rm -f /app/publish/*.pdb /app/publish/*.lic /app/publish/web.config && \
    find /app/publish -maxdepth 1 -name '*.json' ! -name '*.runtimeconfig.json' ! -name '*.deps.json' -delete && \
    rm -rf /app/publish/deploy /app/publish/Tests /app/publish/coverage-reports /app/publish/sonar-reports /app/publish/.claude

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

# Update system packages and install libgdiplus
RUN apt-get update && \
    apt-get upgrade -y && \
    apt-get install -y --no-install-recommends libgdiplus libc6-dev && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Create non-root user for security
RUN groupadd -r aspose && useradd -r -g aspose aspose

# Copy published files
COPY --from=build /app/publish .

# Create directories for license and documents
RUN mkdir -p /app/license /app/documents && \
    chown -R aspose:aspose /app

# Switch to non-root user
USER aspose

# Environment variables
ENV DOTNET_RUNNING_IN_CONTAINER=true

# Entry point (all configuration via environment variables)
ENTRYPOINT ["dotnet", "AsposeMcpServer.dll"]
