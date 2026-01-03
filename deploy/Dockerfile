# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Version from build arg (passed from CI/CD, may include 'v' prefix)
ARG VERSION=1.0.0

# Copy project file and restore dependencies
COPY AsposeMcpServer.csproj .
RUN dotnet restore AsposeMcpServer.csproj

# Copy source code and build
# Remove 'v' prefix from version if present (e.g., v1.0.31 -> 1.0.31)
COPY . .
RUN VERSION_NUM=$(echo "${VERSION}" | sed 's/^v//') && \
    dotnet publish AsposeMcpServer.csproj -c Release -o /app/publish --no-restore -p:Version=${VERSION_NUM}

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

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
