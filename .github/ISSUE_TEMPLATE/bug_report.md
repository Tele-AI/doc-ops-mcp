---
name: Bug report
about: Create a report to help us improve
title: '[BUG] '
labels: bug
assignees: ''

---

**Describe the bug**
A clear and concise description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Install the MCP server with '...'
2. Configure with '...'
3. Run command '...'
4. See error

**Expected behavior**
A clear and concise description of what you expected to happen.

**Actual behavior**
A clear and concise description of what actually happened.

**Screenshots**
If applicable, add screenshots to help explain your problem.

**Environment (please complete the following information):**
 - OS: [e.g. macOS, Windows, Linux]
 - OS Version: [e.g. macOS 14.0, Windows 11, Ubuntu 22.04]
 - Node.js Version: [e.g. 18.17.0]
 - Package Manager: [e.g. npm, pnpm, yarn]
 - MCP Client: [e.g. Claude Desktop, Dive Desktop, custom client]

**Configuration**
```json
// Please provide your MCP server configuration (remove sensitive data)
{
  "mcpServers": {
    "mcp-doc-forge": {
      "command": "...",
      "args": ["..."],
      "env": {
        // your environment variables
      }
    }
  }
}
```

**Logs**
Please provide relevant logs from:
- MCP client logs
- Server output (run with debug mode if possible)
- Any error messages

**Sample Files**
If the issue is related to specific document types, please provide:
- Sample files that reproduce the issue (remove sensitive content)
- File format and size information

**Additional context**
Add any other context about the problem here.