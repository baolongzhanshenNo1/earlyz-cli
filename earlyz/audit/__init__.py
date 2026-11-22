"""早睡 (EarlyZ) package-level audit tools."""

from .ooxml_inventory import build_matrix, scan_and_build, scan_package, write_report

__all__ = ["build_matrix", "scan_and_build", "scan_package", "write_report"]
