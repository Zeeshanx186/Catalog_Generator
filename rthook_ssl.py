"""
Runtime hook for PyInstaller: fix SSL certificate paths for requests / httpx / curl_cffi.

PyInstaller doesn't bundle certifi's CA bundle automatically. Without this hook,
every HTTPS call in the frozen app throws an SSLError that is silently swallowed
by the broad `except Exception` handlers, making ALL searches return zero results.

This runs BEFORE any application code.
"""
import os
import sys

if getattr(sys, "frozen", False):
    try:
        import certifi
        ca_bundle = certifi.where()
        os.environ["SSL_CERT_FILE"] = ca_bundle
        os.environ["REQUESTS_CA_BUNDLE"] = ca_bundle
        # Also covers httpx and curl_cffi which ddgs>=6.x uses internally
        os.environ["CURL_CA_BUNDLE"] = ca_bundle
    except Exception:
        pass  # If certifi itself is missing, fail loudly later
