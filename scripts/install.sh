#!/usr/bin/env bash
# Microsoft MCP Server - Interactive Installer
# Version: 1.0.0
# Repository: https://github.com/marc-hanheide/microsoft-mcp
# License: MIT
# Requires: Bash 4.0+, Python 3.10+, uv or pip, jq (optional)
#
# Usage: curl -fsSL https://raw.githubusercontent.com/marc-hanheide/microsoft-mcp/main/scripts/install.sh | bash
#    or: ./scripts/install.sh [OPTIONS]

set -Eeuo pipefail
IFS=$'\n\t'

# Temporary files tracking for cleanup
declare -a TEMP_FILES=()

# Cleanup trap for temporary files
cleanup() {
    for f in "${TEMP_FILES[@]:-}"; do
        rm -f -- "$f" 2>/dev/null || true
    done
}
trap cleanup EXIT

# Error trap for debugging
trap 'echo "Error at line $LINENO: exit code $?" >&2' ERR

# Check Bash version (requires 4.0+)
if ((BASH_VERSINFO[0] < 4)); then
    echo "Error: This script requires Bash 4.0 or higher" >&2
    echo "Current version: $BASH_VERSION" >&2
    exit 1
fi

# Show usage information
show_usage() {
    cat <<-'USAGE'
	Microsoft MCP Server - Interactive Installer

	USAGE:
	    ./scripts/install.sh [OPTIONS]

	OPTIONS:
	    -h, --help       Show this help message
	    -n, --dry-run    Show what would be done without making changes
	    --version        Show script version

	EXAMPLES:
	    # Interactive installation
	    ./scripts/install.sh

	    # Dry run to see what would happen
	    ./scripts/install.sh --dry-run

	    # Piped installation
	    curl -fsSL https://raw.githubusercontent.com/.../install.sh | bash

	REQUIREMENTS:
	    - Bash 4.0+
	    - Python 3.10+
	    - uv or pip
	    - jq (optional, for config merging)
	USAGE
}

# Parse command line arguments
DRY_RUN=false
while [[ $# -gt 0 ]]; do
    case "$1" in
        -h|--help)
            show_usage
            exit 0
            ;;
        -n|--dry-run)
            DRY_RUN=true
            shift
            ;;
        --version)
            echo "Microsoft MCP Server Installer v1.0.0"
            exit 0
            ;;
        *)
            echo "Unknown option: $1" >&2
            show_usage
            exit 1
            ;;
    esac
done

echo "======================================="
echo "  Microsoft MCP Server Installer"
echo "======================================="
echo ""

if $DRY_RUN; then
    echo "[DRY RUN MODE - No changes will be made]"
    echo ""
fi

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
BOLD='\033[1m'
NC='\033[0m' # No Color

# Installation state
INSTALL_CLAUDE_CODE=false
INSTALL_CURSOR=false
INSTALL_CLAUDE_DESKTOP=false
AUTH_METHOD="msal"  # Default to MSAL (easier, no Azure app registration)
MICROSOFT_MCP_CLIENT_ID=""

# Helper functions
log_info() {
    echo -e "${BLUE}ℹ️  $1${NC}"
}

log_success() {
    echo -e "${GREEN}✅ $1${NC}"
}

log_warning() {
    echo -e "${YELLOW}⚠️  $1${NC}"
}

log_error() {
    echo -e "${RED}❌ $1${NC}"
}

log_header() {
    echo ""
    echo -e "${BOLD}${CYAN}$1${NC}"
    echo "---------------------------------------"
}

# Check if command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Validate Azure client ID format (UUID)
validate_client_id() {
    local id="$1"
    if [[ ! "$id" =~ ^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$ ]]; then
        return 1
    fi
    return 0
}

# Safe directory creation with symlink check
create_config_dir() {
    local dir="$1"

    if [[ -L "$dir" ]]; then
        log_error "Config directory is a symlink: $dir"
        return 1
    fi

    if [[ -e "$dir" && ! -d "$dir" ]]; then
        log_error "Config path exists but is not a directory: $dir"
        return 1
    fi

    if $DRY_RUN; then
        log_info "[DRY RUN] Would create directory: $dir"
        return 0
    fi

    if ! mkdir -p -- "$dir"; then
        log_error "Failed to create directory: $dir"
        return 1
    fi

    # Set restrictive permissions
    chmod 700 -- "$dir" 2>/dev/null || true
    return 0
}

# Create tracked temporary file
create_temp_file() {
    local temp_file
    temp_file=$(mktemp) || {
        log_error "Failed to create temporary file"
        return 1
    }
    TEMP_FILES+=("$temp_file")
    echo "$temp_file"
}

# Detect installation source (local clone vs remote)
# Returns local path if running from a git repo, otherwise GitHub URL
get_install_source() {
    local script_dir
    script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
    local repo_root="${script_dir%/scripts}"

    # Check if we're in a git repo with pyproject.toml (local development)
    if [[ -f "$repo_root/pyproject.toml" ]] && [[ -d "$repo_root/.git" ]]; then
        echo "$repo_root"
    else
        echo "git+https://github.com/marc-hanheide/microsoft-mcp.git"
    fi
}

# Detect OS
detect_os() {
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "macos"
    elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
        echo "linux"
    elif [[ "$OSTYPE" == "msys" ]] || [[ "$OSTYPE" == "cygwin" ]]; then
        echo "windows"
    else
        echo "unknown"
    fi
}

# Get Claude Desktop config directory
get_claude_desktop_config_dir() {
    local os=$(detect_os)
    case "$os" in
        macos)
            echo "$HOME/Library/Application Support/Claude"
            ;;
        linux)
            echo "$HOME/.config/Claude"
            ;;
        windows)
            echo "$APPDATA/Claude"
            ;;
        *)
            echo ""
            ;;
    esac
}

# Check prerequisites
check_prerequisites() {
    log_header "Checking Prerequisites"

    local os=$(detect_os)
    log_info "Detected OS: $os"

    # Check Python
    if command_exists python3; then
        local python_version=$(python3 --version 2>&1 | cut -d' ' -f2)
        log_success "Python $python_version found"
    elif command_exists python; then
        local python_version=$(python --version 2>&1 | cut -d' ' -f2)
        log_success "Python $python_version found"
    else
        log_error "Python not found. Please install Python 3.10+ first."
        log_info "Visit: https://www.python.org/downloads/"
        exit 1
    fi

    # Check uv (preferred) or pip
    if command_exists uv; then
        log_success "uv found (recommended)"
    elif command_exists uvx; then
        log_success "uvx found"
    elif command_exists pip; then
        log_warning "pip found, but uv is recommended for better performance"
        log_info "Install uv: curl -LsSf https://astral.sh/uv/install.sh | sh"
    else
        log_error "Neither uv nor pip found. Please install uv first."
        log_info "Install uv: curl -LsSf https://astral.sh/uv/install.sh | sh"
        exit 1
    fi

    # Check jq (optional but recommended)
    if command_exists jq; then
        log_success "jq found (for config merging)"
    else
        log_warning "jq not found - config merging will overwrite existing configs"
        log_info "Install jq: brew install jq (macOS) or apt install jq (Linux)"
    fi

    # Check Claude CLI (optional)
    if command_exists claude; then
        log_success "Claude CLI found"
    else
        log_warning "Claude CLI not found - Claude Code installation will be skipped"
        log_info "Install: npm install -g @anthropic-ai/claude-code"
    fi
}

# Interactive target selection
select_install_targets() {
    log_header "Select Installation Targets"

    echo "Which clients would you like to configure?"
    echo "(Enter numbers separated by spaces, or 'a' for all)"
    echo ""

    local has_claude_cli=false
    command_exists claude && has_claude_cli=true

    if $has_claude_cli; then
        echo "  1) Claude Code (CLI)"
    else
        echo -e "  ${YELLOW}1) Claude Code (CLI) - requires Claude CLI${NC}"
    fi
    echo "  2) Cursor IDE"
    echo "  3) Claude Desktop"
    echo ""
    echo "  a) All available"
    echo "  q) Quit"
    echo ""

    read -r -p "Your selection: " selection

    case "$selection" in
        *q*|*Q*)
            echo "Installation cancelled."
            exit 0
            ;;
        *a*|*A*)
            $has_claude_cli && INSTALL_CLAUDE_CODE=true
            INSTALL_CURSOR=true
            INSTALL_CLAUDE_DESKTOP=true
            ;;
        *)
            [[ "$selection" == *"1"* ]] && $has_claude_cli && INSTALL_CLAUDE_CODE=true
            [[ "$selection" == *"2"* ]] && INSTALL_CURSOR=true
            [[ "$selection" == *"3"* ]] && INSTALL_CLAUDE_DESKTOP=true
            ;;
    esac

    # Verify at least one target selected
    if ! $INSTALL_CLAUDE_CODE && ! $INSTALL_CURSOR && ! $INSTALL_CLAUDE_DESKTOP; then
        log_error "No valid targets selected."
        exit 1
    fi

    echo ""
    log_info "Selected targets:"
    $INSTALL_CLAUDE_CODE && echo "  - Claude Code (CLI)"
    $INSTALL_CURSOR && echo "  - Cursor IDE"
    $INSTALL_CLAUDE_DESKTOP && echo "  - Claude Desktop"
}

# Interactive auth method selection
select_auth_method() {
    log_header "Select Authentication Method"

    echo "Microsoft MCP supports two authentication methods:"
    echo ""
    echo -e "  ${GREEN}1) MSAL Device Code Flow (Recommended)${NC}"
    echo "     - Works in CLI/headless environments"
    echo "     - No Azure app registration required"
    echo "     - Uses Microsoft Office client ID"
    echo "     - Displays code to enter at microsoft.com/devicelogin"
    echo ""
    echo "  2) Azure SDK Browser Flow"
    echo "     - Opens browser for sign-in"
    echo "     - Requires Azure app registration"
    echo "     - Platform-specific secure token storage"
    echo ""

    read -r -p "Your selection [1]: " auth_selection

    case "$auth_selection" in
        2)
            AUTH_METHOD="azure"
            log_info "Selected: Azure SDK Browser Flow"

            # Get Azure app ID
            echo ""
            log_info "Azure SDK requires an Azure AD Application ID."
            echo -e "${YELLOW}Get yours from: https://portal.azure.com → Microsoft Entra ID → App registrations${NC}"
            echo ""
            read -r -p "Enter your Azure Application (Client) ID: " client_id

            if [ -z "$client_id" ]; then
                log_error "Azure Application ID is required for Azure SDK auth."
                log_info "Falling back to MSAL authentication."
                AUTH_METHOD="msal"
            elif ! validate_client_id "$client_id"; then
                log_error "Invalid Azure Application ID format (expected UUID like xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)"
                log_info "Falling back to MSAL authentication."
                AUTH_METHOD="msal"
            else
                MICROSOFT_MCP_CLIENT_ID="$client_id"
                log_success "Azure Application ID saved"
            fi
            ;;
        *)
            AUTH_METHOD="msal"
            log_info "Selected: MSAL Device Code Flow (no setup required)"
            ;;
    esac
}

# Backup existing config file
backup_config() {
    local config_file="$1"

    if [[ -f "$config_file" ]]; then
        # Check read permission
        if [[ ! -r "$config_file" ]]; then
            log_warning "Cannot read $config_file (permission denied), skipping backup"
            return 0
        fi

        local backup_dir
        backup_dir=$(dirname -- "$config_file")

        # Check write permission for backup
        if [[ ! -w "$backup_dir" ]]; then
            log_warning "Cannot write to $backup_dir, skipping backup"
            return 0
        fi

        if $DRY_RUN; then
            log_info "[DRY RUN] Would backup $config_file"
            return 0
        fi

        local backup_file="${config_file}.backup.$(date +%Y%m%d_%H%M%S)"
        if cp -- "$config_file" "$backup_file"; then
            log_success "Backed up existing config to: $(basename -- "$backup_file")"
        else
            log_warning "Failed to create backup, continuing anyway"
        fi
    fi
}

# Build the MCP server JSON config
build_server_config() {
    local config=""
    local install_source
    install_source=$(get_install_source)

    # Log the installation source for visibility (to stderr to not pollute JSON output)
    if [[ "$install_source" != git+* ]]; then
        log_info "Using local installation source: $install_source" >&2
    fi

    if [ "$AUTH_METHOD" == "msal" ]; then
        config=$(cat <<EOF
{
    "command": "uvx",
    "args": ["--python", "3.13", "--from", "$install_source", "microsoft-mcp"],
    "env": {
        "MICROSOFT_MCP_AUTH_METHOD": "msal"
    }
}
EOF
)
    else
        config=$(cat <<EOF
{
    "command": "uvx",
    "args": ["--python", "3.13", "--from", "$install_source", "microsoft-mcp"],
    "env": {
        "MICROSOFT_MCP_CLIENT_ID": "$MICROSOFT_MCP_CLIENT_ID"
    }
}
EOF
)
    fi

    echo "$config"
}

# Install for Claude Code (CLI)
install_claude_code() {
    log_header "Configuring Claude Code (CLI)"

    if ! command_exists claude; then
        log_error "Claude CLI not found. Skipping."
        return 1
    fi

    local server_config
    server_config=$(build_server_config) || {
        log_error "Failed to build server config"
        return 1
    }

    # Check if already configured
    if claude mcp list 2>/dev/null | grep -q "microsoft-mcp"; then
        log_info "microsoft-mcp already configured. Updating..."
        if ! $DRY_RUN; then
            claude mcp remove microsoft-mcp -s user 2>/dev/null || true
        fi
    fi

    # Add the MCP server
    log_info "Adding microsoft-mcp to Claude Code..."

    if $DRY_RUN; then
        log_info "[DRY RUN] Would add microsoft-mcp server config"
        log_success "Claude Code configured successfully (dry run)"
        return 0
    fi

    if claude mcp add-json microsoft-mcp "$server_config" -s user; then
        log_success "Claude Code configured successfully"
    else
        log_error "Failed to configure Claude Code"
        return 1
    fi
}

# Install for Cursor IDE
install_cursor() {
    log_header "Configuring Cursor IDE"

    local config_dir="$HOME/.cursor"
    local config_file="$config_dir/mcp.json"

    # Create directory if needed (with symlink check)
    if ! create_config_dir "$config_dir"; then
        return 1
    fi

    # Backup existing config
    backup_config "$config_file"

    local server_config
    server_config=$(build_server_config) || {
        log_error "Failed to build server config"
        return 1
    }

    if $DRY_RUN; then
        log_info "[DRY RUN] Would configure Cursor at: $config_file"
        log_success "Cursor configured (dry run)"
        return 0
    fi

    if [ -f "$config_file" ] && command_exists jq; then
        # Merge with existing config using jq
        log_info "Merging with existing Cursor config..."

        local temp_file
        temp_file=$(create_temp_file) || return 1

        if jq -e '.mcpServers' "$config_file" >/dev/null 2>&1; then
            if ! jq --argjson microsoft "$server_config" \
                '.mcpServers["microsoft-mcp"] = $microsoft' \
                "$config_file" > "$temp_file"; then
                log_error "Failed to merge config with jq"
                return 1
            fi
        else
            if ! jq --argjson microsoft "$server_config" \
                '. + {"mcpServers": {"microsoft-mcp": $microsoft}}' \
                "$config_file" > "$temp_file"; then
                log_error "Failed to create config with jq"
                return 1
            fi
        fi

        # Validate output is valid JSON
        if ! jq empty "$temp_file" 2>/dev/null; then
            log_error "Generated invalid JSON"
            return 1
        fi

        mv -- "$temp_file" "$config_file"
        log_success "Cursor configuration merged"
    else
        # Create new config
        log_info "Creating new Cursor config..."

        # Build config with jq if available for proper escaping
        if command_exists jq; then
            jq -n --argjson microsoft "$server_config" \
                '{"mcpServers": {"microsoft-mcp": $microsoft}}' > "$config_file" || {
                log_error "Failed to create config file"
                return 1
            }
        else
            cat > "$config_file" << EOF
{
  "mcpServers": {
    "microsoft-mcp": $server_config
  }
}
EOF
        fi
        log_success "Cursor configuration created"
    fi

    log_success "Cursor configured at: $config_file"
}

# Install for Claude Desktop
install_claude_desktop() {
    log_header "Configuring Claude Desktop"

    local config_dir
    config_dir=$(get_claude_desktop_config_dir)

    if [ -z "$config_dir" ]; then
        log_error "Could not determine Claude Desktop config directory for this OS."
        return 1
    fi

    local config_file="$config_dir/claude_desktop_config.json"

    # Create directory if needed (with symlink check)
    if ! create_config_dir "$config_dir"; then
        return 1
    fi

    # Backup existing config
    backup_config "$config_file"

    local server_config
    server_config=$(build_server_config) || {
        log_error "Failed to build server config"
        return 1
    }

    if $DRY_RUN; then
        log_info "[DRY RUN] Would configure Claude Desktop at: $config_file"
        log_success "Claude Desktop configured (dry run)"
        return 0
    fi

    if [ -f "$config_file" ] && command_exists jq; then
        # Merge with existing config using jq
        log_info "Merging with existing Claude Desktop config..."

        local temp_file
        temp_file=$(create_temp_file) || return 1

        if jq -e '.mcpServers' "$config_file" >/dev/null 2>&1; then
            if ! jq --argjson microsoft "$server_config" \
                '.mcpServers["microsoft-mcp"] = $microsoft' \
                "$config_file" > "$temp_file"; then
                log_error "Failed to merge config with jq"
                return 1
            fi
        else
            if ! jq --argjson microsoft "$server_config" \
                '. + {"mcpServers": {"microsoft-mcp": $microsoft}}' \
                "$config_file" > "$temp_file"; then
                log_error "Failed to create config with jq"
                return 1
            fi
        fi

        # Validate output is valid JSON
        if ! jq empty "$temp_file" 2>/dev/null; then
            log_error "Generated invalid JSON"
            return 1
        fi

        mv -- "$temp_file" "$config_file"
        log_success "Claude Desktop configuration merged"
    else
        # Create new config
        log_info "Creating new Claude Desktop config..."

        # Build config with jq if available for proper escaping
        if command_exists jq; then
            jq -n --argjson microsoft "$server_config" \
                '{"mcpServers": {"microsoft-mcp": $microsoft}}' > "$config_file" || {
                log_error "Failed to create config file"
                return 1
            }
        else
            cat > "$config_file" << EOF
{
  "mcpServers": {
    "microsoft-mcp": $server_config
  }
}
EOF
        fi
        log_success "Claude Desktop configuration created"
    fi

    log_success "Claude Desktop configured at: $config_file"
}

# Run initial authentication
run_authentication() {
    log_header "Initial Authentication"

    local install_source
    install_source=$(get_install_source)

    echo "Would you like to authenticate now?"
    echo "This will allow you to verify the setup is working."
    echo ""
    read -r -p "Run authentication? [Y/n]: " run_auth

    case "$run_auth" in
        [nN]*)
            log_info "Skipping authentication. You can run it later with:"
            if [ "$AUTH_METHOD" == "msal" ]; then
                echo "  MICROSOFT_MCP_AUTH_METHOD=msal uvx --python 3.13 --from $install_source python -c 'from microsoft_mcp.auth_msal import MSALRefreshTokenAuth; a = MSALRefreshTokenAuth(); a.authenticate()'"
            else
                echo "  MICROSOFT_MCP_CLIENT_ID=$MICROSOFT_MCP_CLIENT_ID uvx --python 3.13 --from $install_source python -c 'from microsoft_mcp.auth import AzureAuthentication; a = AzureAuthentication(); a.authenticate()'"
            fi
            return
            ;;
    esac

    log_info "Starting authentication..."
    echo ""

    if [ "$AUTH_METHOD" == "msal" ]; then
        log_info "MSAL Device Code Flow: You'll see a code to enter at microsoft.com/devicelogin"
        echo ""
        MICROSOFT_MCP_AUTH_METHOD=msal uvx --python 3.13 --from "$install_source" python -c "
from microsoft_mcp.auth_msal import MSALRefreshTokenAuth
auth = MSALRefreshTokenAuth()
result = auth.authenticate()
print('Authentication successful!')
print(f'Logged in as: {result.get(\"username\", \"unknown\")}')
" || {
            log_warning "Authentication failed or was cancelled."
            log_info "You can try again later."
        }
    else
        log_info "Azure SDK: A browser window will open for sign-in"
        echo ""
        MICROSOFT_MCP_CLIENT_ID="$MICROSOFT_MCP_CLIENT_ID" uvx --python 3.13 --from "$install_source" python -c "
from microsoft_mcp.auth import AzureAuthentication
auth = AzureAuthentication()
auth.authenticate()
print('Authentication successful!')
" || {
            log_warning "Authentication failed or was cancelled."
            log_info "You can try again later."
        }
    fi
}

# Show completion message and next steps
show_completion() {
    echo ""
    echo "======================================="
    echo -e "${GREEN}${BOLD}  Installation Complete!${NC}"
    echo "======================================="
    echo ""

    echo "Configured targets:"
    $INSTALL_CLAUDE_CODE && echo "  ✅ Claude Code (CLI)"
    $INSTALL_CURSOR && echo "  ✅ Cursor IDE"
    $INSTALL_CLAUDE_DESKTOP && echo "  ✅ Claude Desktop"
    echo ""

    echo "Authentication method: $AUTH_METHOD"
    echo ""

    echo -e "${BOLD}Next Steps:${NC}"
    echo ""

    if $INSTALL_CURSOR || $INSTALL_CLAUDE_DESKTOP; then
        echo "1. Restart your applications:"
        $INSTALL_CURSOR && echo "   - Quit and reopen Cursor IDE"
        $INSTALL_CLAUDE_DESKTOP && echo "   - Quit and reopen Claude Desktop"
        echo ""
    fi

    echo "2. Try it out! Example prompts:"
    echo "   - \"Read my latest emails\""
    echo "   - \"Show my calendar for this week\""
    echo "   - \"Search for files about project alpha\""
    echo "   - \"List my OneDrive files\""
    echo ""

    echo -e "${BOLD}Documentation:${NC}"
    echo "   https://github.com/marc-hanheide/microsoft-mcp"
    echo ""

    echo -e "${BOLD}Troubleshooting:${NC}"
    echo "   - Re-run authentication: ./scripts/install.sh"
    echo "   - Check logs: Look for errors in the client's MCP logs"
    echo "   - Issues: https://github.com/marc-hanheide/microsoft-mcp/issues"
    echo ""
}

# Main installation flow
main() {
    check_prerequisites
    select_install_targets
    select_auth_method

    echo ""
    log_header "Installing Microsoft MCP Server"

    # Run installations with error tracking
    local install_failed=false

    if $INSTALL_CLAUDE_CODE; then
        if ! install_claude_code; then
            install_failed=true
        fi
    fi

    if $INSTALL_CURSOR; then
        if ! install_cursor; then
            install_failed=true
        fi
    fi

    if $INSTALL_CLAUDE_DESKTOP; then
        if ! install_claude_desktop; then
            install_failed=true
        fi
    fi

    if $install_failed; then
        log_warning "Some installations failed. Check messages above."
    fi

    # Offer to run authentication (skip in dry run mode)
    if ! $DRY_RUN; then
        run_authentication
    else
        log_info "[DRY RUN] Skipping authentication"
    fi

    # Show completion
    show_completion
}

# Run main function
main "$@"
