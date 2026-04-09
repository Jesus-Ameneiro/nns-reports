"""
Bentley — NNS Evidence Report Generator (stub)
Business logic to be implemented in bentley/processor.py and bentley/report_writer.py.
This module exposes a single function:  render()
Called by the root app.py inside the Bentley tab.
"""

import streamlit as st


def render():
    """Render the Bentley Evidence Report Generator UI inside its tab."""
    st.markdown(
        '<div class="alert alert-warn" style="margin-top:1.5rem;">'
        '🚧 &nbsp; The Bentley NNS Evidence Report Generator is under development.'
        '</div>',
        unsafe_allow_html=True,
    )
