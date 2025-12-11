#!/bin/bash
# Aspose MCP Server - Cross-platform Build Script (Linux/macOS)
# 跨平台构建脚本

set -e

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
GRAY='\033[0;37m'
NC='\033[0m' # No Color

echo -e "${CYAN}=== Aspose MCP Server - Cross-platform Build ===${NC}"
echo ""

# Parse arguments
CLEAN=false
WINDOWS=false
LINUX=false
MACOS=false
MACOS_X64=false
MACOS_ARM64=false
ALL=false
PLATFORM_SPECIFIC=""

for arg in "$@"; do
    case $arg in
        --clean)
            CLEAN=true
            ;;
        --windows)
            WINDOWS=true
            ;;
        --linux)
            LINUX=true
            ;;
        --macos)
            MACOS=true
            ;;
        linux-x64)
            LINUX=true
            PLATFORM_SPECIFIC="linux-x64"
            ;;
        macos-x64)
            MACOS_X64=true
            PLATFORM_SPECIFIC="macos-x64"
            ;;
        macos-arm64)
            MACOS_ARM64=true
            PLATFORM_SPECIFIC="macos-arm64"
            ;;
        --all)
            ALL=true
            ;;
        --help|-h)
            echo -e "${YELLOW}Usage:${NC}"
            echo -e "  ${GRAY}./publish.sh --windows    # Build for Windows${NC}"
            echo -e "  ${GRAY}./publish.sh --linux      # Build for Linux${NC}"
            echo -e "  ${GRAY}./publish.sh --macos      # Build for macOS (Intel + ARM)${NC}"
            echo -e "  ${GRAY}./publish.sh macos-x64    # Build for macOS Intel only${NC}"
            echo -e "  ${GRAY}./publish.sh macos-arm64  # Build for macOS ARM only${NC}"
            echo -e "  ${GRAY}./publish.sh --all        # Build for all platforms${NC}"
            echo -e "  ${GRAY}./publish.sh --clean      # Clean before build${NC}"
            echo ""
            echo -e "${YELLOW}Example:${NC}"
            echo -e "  ${GRAY}./publish.sh --all --clean${NC}"
            exit 0
            ;;
    esac
done

# Clean if requested
if [ "$CLEAN" = true ] || [ "$ALL" = true ]; then
    echo -e "${YELLOW}Cleaning output directory...${NC}"
    rm -rf publish
fi

# Create output directory
mkdir -p publish

build_platform() {
    local runtime=$1
    local platform=$2
    
    echo -e "${GREEN}Building for $platform ($runtime)...${NC}"
    
    local output_path="publish/$platform"
    
    # Get version from Git tag if available, otherwise use default
    if [ -n "$VERSION" ]; then
        version="$VERSION"
    else
        git_tag=$(git describe --tags --abbrev=0 2>/dev/null || echo "v1.0.0")
        version=$(echo "$git_tag" | sed 's/^v//')
    fi
    
    dotnet publish \
        --configuration Release \
        --runtime "$runtime" \
        --self-contained true \
        --output "$output_path" \
        -p:Version="$version" \
        -p:PublishSingleFile=true \
        -p:PublishTrimmed=false \
        -p:IncludeNativeLibrariesForSelfExtract=true \
        --nologo \
        --verbosity quiet
    
    if [ $? -eq 0 ]; then
        echo -e "  ${GREEN}✓ Build successful: $output_path${NC}"
        
        # Copy license file
        if [ -f "Aspose.Total.lic" ]; then
            cp "Aspose.Total.lic" "$output_path/"
            echo -e "  ${GREEN}✓ License file copied${NC}"
        fi
        
        # Make executable (for Linux/macOS)
        if [ "$runtime" != "win-x64" ]; then
            chmod +x "$output_path/AsposeMcpServer"
            echo -e "  ${GREEN}✓ Set executable permission${NC}"
        fi
        
        # Get directory size
        size=$(du -sm "$output_path" | cut -f1)
        echo -e "  ${GRAY}Size: $size MB${NC}"
    else
        echo -e "  ${RED}✗ Build failed${NC}"
    fi
    echo ""
}

# Build for selected platforms
if [ "$ALL" = true ] || [ "$WINDOWS" = true ]; then
    build_platform "win-x64" "windows-x64"
fi

if [ "$ALL" = true ] || [ "$LINUX" = true ]; then
    build_platform "linux-x64" "linux-x64"
fi

if [ "$ALL" = true ] || [ "$MACOS" = true ]; then
    build_platform "osx-x64" "macos-x64"
    build_platform "osx-arm64" "macos-arm64"
elif [ "$MACOS_X64" = true ]; then
    build_platform "osx-x64" "macos-x64"
elif [ "$MACOS_ARM64" = true ]; then
    build_platform "osx-arm64" "macos-arm64"
fi

# If no platform specified, show help
if [ "$WINDOWS" = false ] && [ "$LINUX" = false ] && [ "$MACOS" = false ] && [ "$MACOS_X64" = false ] && [ "$MACOS_ARM64" = false ] && [ "$ALL" = false ]; then
    echo -e "${YELLOW}Usage:${NC}"
    echo -e "  ${GRAY}./publish.sh --windows    # Build for Windows${NC}"
    echo -e "  ${GRAY}./publish.sh --linux      # Build for Linux${NC}"
    echo -e "  ${GRAY}./publish.sh --macos      # Build for macOS (Intel + ARM)${NC}"
    echo -e "  ${GRAY}./publish.sh --all        # Build for all platforms${NC}"
    echo -e "  ${GRAY}./publish.sh --clean      # Clean before build${NC}"
    echo ""
    echo -e "${YELLOW}Example:${NC}"
    echo -e "  ${GRAY}./publish.sh --all --clean${NC}"
    exit 0
fi

echo -e "${CYAN}=== Build Complete ===${NC}"
echo ""
echo -e "${GREEN}Output directory: $(pwd)/publish${NC}"
echo ""
echo -e "${YELLOW}Usage examples:${NC}"
echo ""
echo -e "${CYAN}Windows:${NC}"
echo -e '  ${GRAY}"C:\path\to\publish\windows-x64\AsposeMcpServer.exe" --word${NC}'
echo ""
echo -e "${CYAN}Linux/macOS:${NC}"
echo -e '  ${GRAY}/path/to/publish/linux-x64/AsposeMcpServer --word${NC}'
echo ""
echo -e "${CYAN}Claude Desktop config.json:${NC}"
echo -e "${GRAY}{"
echo -e '  "mcpServers": {'
echo -e '    "aspose-word": {'
echo -e '      "command": "/path/to/AsposeMcpServer",'
echo -e '      "args": ["--word"]'
echo -e '    }'
echo -e '  }'
echo -e "}${NC}"

