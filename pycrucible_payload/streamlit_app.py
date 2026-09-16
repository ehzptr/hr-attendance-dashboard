"""
Streamlit entrypoint.

Run with:
    streamlit run streamlit_app.py

This file only exists so Streamlit has a plain script to execute (Streamlit
scripts can't use package-relative imports). All real logic lives in the
``hrdash`` package.
"""

from hrdash.app import main

main()
