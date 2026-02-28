"""
Actions package for the Word MCP SSE Server.

Each module in this package contains one async action function ending in _action.
These are auto-discovered and registered as MCP tools at startup.

Available actions:
- list_templates_action     → list_templates_tool
- list_projects_action      → list_projects_tool
- read_template_placeholders_action → read_template_placeholders_tool
- fill_template_action      → fill_template_tool
- get_context_action        → get_context_tool
- edit_document_action      → edit_document_tool
- status_action             → status_tool
"""