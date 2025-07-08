"""
Main entry point for the Spire.Doc MCP Server.
"""
import asyncio
import os
from .server import run_server

def main():
    """Start the SpireDoc MCP server."""
    try:
        print("SpireDoc MCP Server")
        print("---------------")
        print("Starting server... Press Ctrl+C to exit")       
        asyncio.run(run_server())
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Server stopped.")

if __name__ == "__main__":
    main() 