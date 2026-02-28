"""
Word MCP SSE Server entry point.
"""

import argparse
import logging
from typing import cast

import uvicorn

from src.config import load_config
from src.mcp_tools import MCPServer, register_tools


def setup_logging(
    log_level: str = "INFO", file_logging: bool = False, logs_dir: str = "logs"
) -> logging.Logger:
    """Configure application logging with optional file logging."""
    from pathlib import Path

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    logger = logging.getLogger("word-mcp-server")
    logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))

    # Clear any existing handlers
    logger.handlers = []

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler if enabled
    if file_logging:
        logs_path = Path(logs_dir)
        logs_path.mkdir(exist_ok=True)
        log_file_path = logs_path / "word-mcp-server.log"
        file_handler = logging.FileHandler(log_file_path)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.info(f"File logging enabled: {log_file_path}")

    logger.info("Logging configured")
    return logger


def main():
    """Main entry point for the application."""
    try:
        # Parse command line arguments
        parser = argparse.ArgumentParser()
        parser.add_argument("--host", default="0.0.0.0")
        parser.add_argument("--port", type=int, default=8080)
        parser.add_argument("--log-level", default=None, help="Override log level")
        args = parser.parse_args()

        # Load configuration
        config = load_config()

        # Set up logging
        log_level = args.log_level or config.LOG_LEVEL
        logger = setup_logging(log_level, config.FILE_LOGGING, config.LOGS_DIR)

        if log_level.upper() == "DEBUG":
            logger.warning("DEBUG logging enabled - ensure no secrets are logged")

        if args.log_level:
            logger.info(f"Log level overridden to {args.log_level.upper()}")

        # Initialize MCP server
        logger.info("Initializing Word MCP server")
        mcp_server = MCPServer(api_key=cast(str, config.MCP_SERVER_AUTH_KEY))

        # Register tools
        logger.info("Registering Word MCP tools")
        register_tools(mcp_server=mcp_server)

        # Create and run app
        logger.info(f"Starting Word MCP server on http://{args.host}:{args.port}")
        app = mcp_server.create_app(debug=True)

        # Start server
        logger.info("Starting Uvicorn server")
        uvicorn.run(app, host=args.host, port=args.port, log_level="info")

    except Exception as e:
        if "logger" in locals():
            logger.critical(f"Failed to start application: {str(e)}", exc_info=True)
        else:
            logging.critical(f"Failed to start application: {str(e)}", exc_info=True)
        raise


if __name__ == "__main__":
    main()