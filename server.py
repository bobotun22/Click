# server.py — Streamlit wrapper for the Bank Transaction app
import streamlit.components.v1 as components

# Read the HTML file
with open("index.html", "r", encoding="utf-8") as f:
    html_content = f.read()

# Render full-page in an iframe
components.html(
    html_content,
    height=900,
    scrolling=True,
)
