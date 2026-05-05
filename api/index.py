"""
api/index.py — Vercel entrypoint for BRIMI Flask app.

Vercel expects an `app` callable at this module level.
Flask's WSGI app satisfies this — Fluid Compute keeps it warm.
"""
import sys
import os

# Ensure project root is on the path so Flask imports sibling modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app  # noqa: F401 — Vercel needs the Flask app callable
