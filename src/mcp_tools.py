"""
MCP tools registration for the Word MCP SSE Server.
"""

import importlib
import inspect
import logging
import pkgutil
import uuid
from typing import Callable, TypeVar

from mcp.server.fastmcp import FastMCP
from mcp.server.sse import SseServerTransport
from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Mount, Route

from . import actions

# ------------------------------------------------------------
# Central place where *all* server-supplied objects live
DEPENDENCIES: dict[str, object] = {
    # These will be populated by register_tools()
    # append new shared objects here ↓
}

T = TypeVar("T")
logger = logging.getLogger(__name__)


class APIKeyMiddleware(BaseHTTPMiddleware):
    """Middleware for API key authentication."""

    def __init__(self, app, api_key: str):
        super().__init__(app)
        self.api_key = api_key

    async def dispatch(self, request: Request, call_next):
        request_id = str(uuid.uuid4())

        logger.info(f"[{request_id}] {request.method} {request.url.path}")

        # Skip auth for health endpoint
        if request.url.path == "/health":
            return await call_next(request)

        # Check API key
        if request.headers.get("X-API-Key") == self.api_key:
            logger.debug(f"[{request_id}] API key authentication successful")
            response = await call_next(request)
            logger.info(f"[{request_id}] Completed with status {response.status_code}")
            return response
        else:
            logger.warning(f"[{request_id}] Unauthorized: Invalid API key")
            return JSONResponse({"error": "Unauthorized"}, status_code=401)


class MCPServer:
    """Simplified MCP server."""

    def __init__(self, api_key: str, service_name: str = "word-mcp-server"):
        self.api_key = api_key
        self.mcp = FastMCP(service_name)
        logger.info(f"Initialized MCP server: {service_name}")

    def register_tool(self, func: Callable[..., T]) -> Callable[..., T]:
        """Register a function as an MCP tool."""
        logger.info(f"Registering MCP tool: {func.__name__}")
        return self.mcp.tool()(func)

    def create_app(self, debug: bool = False) -> Starlette:
        """Create a Starlette application with MCP server."""
        sse = SseServerTransport("/messages/")

        async def handle_sse(request: Request) -> JSONResponse | None:
            request_id = str(uuid.uuid4())
            logger.info(f"[{request_id}] SSE connection established")

            if request.method in {"HEAD", "OPTIONS"}:
                logger.debug(
                    f"[{request_id}] Non-streaming method {request.method} received – returning 200"
                )
                return JSONResponse({"status": "ok"}, status_code=200)

            try:
                async with sse.connect_sse(
                    request.scope, request.receive, request._send
                ) as (read_stream, write_stream):
                    await self.mcp._mcp_server.run(
                        read_stream,
                        write_stream,
                        self.mcp._mcp_server.create_initialization_options(),
                    )
            except Exception as e:
                logger.error(f"[{request_id}] SSE error: {str(e)}", exc_info=True)
                raise
            finally:
                logger.info(f"[{request_id}] SSE connection closed")

        async def handle_health(request: Request) -> JSONResponse:
            """Health check endpoint."""
            return JSONResponse({
                "status": "healthy",
                "service": "word-mcp-server",
                "version": "1.0.0"
            }, status_code=200)

        # Health endpoint bypasses API key middleware
        health_routes = [Route("/health", endpoint=handle_health)]

        # Protected routes with API key middleware
        protected_middleware = [Middleware(APIKeyMiddleware, api_key=self.api_key)]
        protected_routes = [
            Route("/sse", endpoint=handle_sse),
            Mount("/messages/", app=sse.handle_post_message),
        ]

        app = Starlette(
            debug=debug,
            routes=health_routes + protected_routes,
            middleware=protected_middleware,
        )

        logger.info("Starlette application created")
        return app


def make_wrapper(action_func):
    """Create wrapper that injects only the dependencies the action explicitly asks for."""
    sig = inspect.signature(action_func)
    wanted = {
        name: value
        for name, value in DEPENDENCIES.items()
        if name in sig.parameters
    }

    async def wrapper(**kwargs):
        kwargs.update(wanted)
        return await action_func(**kwargs)

    wrapper.__name__ = action_func.__name__.replace("_action", "_tool")
    wrapper.__doc__ = action_func.__doc__

    # Build a new signature that excludes injected parameters
    params = [
        p for p in sig.parameters.values()
        if p.name not in wanted
    ]
    wrapper.__signature__ = inspect.Signature(
        parameters=params,
        return_annotation=sig.return_annotation,
    )

    # Copy annotations but remove injected parameters
    if hasattr(action_func, "__annotations__"):
        wrapper.__annotations__ = {
            k: v for k, v in action_func.__annotations__.items()
            if k not in wanted
        }

    return wrapper


def register_tools(mcp_server: MCPServer) -> None:
    """Register all MCP tools by auto-discovering action modules."""

    # Populate the dependencies registry with Graph client settings
    from .config import load_config
    settings = load_config()

    DEPENDENCIES.update({
        "azure_tenant_id": settings.AZURE_TENANT_ID,
        "azure_client_id": settings.AZURE_CLIENT_ID,
        "azure_client_secret": settings.AZURE_CLIENT_SECRET,
        "sharepoint_drive_id": settings.SHAREPOINT_DRIVE_ID,      # ← changed
        "sharepoint_template_folder": settings.SHAREPOINT_TEMPLATE_FOLDER,
        "onedrive_user": settings.ONEDRIVE_USER,
        "onedrive_output_folder": settings.ONEDRIVE_OUTPUT_FOLDER,
    })

    logger.info("Starting auto-discovery of action modules")

    # Auto-discover and register all action functions
    for _, module_name, _ in pkgutil.iter_modules(actions.__path__):
        try:
            mod = importlib.import_module(
                f".actions.{module_name}", package=__package__
            )
            logger.debug(f"Loaded action module: {module_name}")

            for name, func in inspect.getmembers(mod, inspect.iscoroutinefunction):
                if name.endswith("_action"):
                    logger.info(f"Registering action: {name}")
                    tool_wrapper = make_wrapper(func)
                    mcp_server.register_tool(tool_wrapper)

        except Exception as e:
            logger.error(
                f"Failed to load action module {module_name}: {str(e)}", exc_info=True
            )
            raise

    logger.info("Action module auto-discovery completed")