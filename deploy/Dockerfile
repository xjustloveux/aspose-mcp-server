# Build stage
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# Copy project file and restore dependencies
COPY AsposeMcpServer.csproj .
RUN dotnet restore

# Copy source code and build
COPY . .
RUN dotnet publish -c Release -o /app/publish --no-restore

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
