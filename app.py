"""Application entry point and backwards-compatible public imports.

The implementation is organized by responsibility in inventory_core,
inventory_service, inventory_web, and inventory_logging.
"""

from inventory_core import *
from inventory_service import InventoryApp
from inventory_web import APP, RequestHandler, choose_directory, main


if __name__ == "__main__":
    main()
