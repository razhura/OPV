# navbar.py
import streamlit as st

def render_navbar():
    st.markdown("""
    <style>
    .navbar {
        display: flex;
        gap: 2rem;
        margin-bottom: 2rem;
    }

    .nav-item {
        position: relative;
        font-size: 18px;
        font-weight: bold;
        color: white;
        background-color: #4A6C6F;
        padding: 0.5rem 1rem;
        border-radius: 10px;
        cursor: pointer;
        transition: all 0.3s ease-in-out;
        width: 60px;
        text-align: center;
        text-decoration: none;
    }

    .nav-item:hover {
        width: 280px;
        background-color: #3C5C5F;
    }

    .nav-item span.full-text {
        opacity: 0;
        transition: opacity 0.3s ease-in-out;
        position: absolute;
        left: 70px;
        top: 50%;
        transform: translateY(-50%);
        white-space: nowrap;
    }

    .nav-item:hover span.full-text {
        opacity: 1;
    }

    .nav-container {
        display: flex;
        justify-content: flex-start;
        margin-top: 1rem;
    }
    </style>

    <div class="nav-container">
        <div class="navbar">
            <a href="/?page=QCA" class="nav-item">
                QCA
                <span class="full-text">Critical Quality Attribute</span>
            </a>
            <a href="/?page=IPC" class="nav-item">
                IPC
                <span class="full-text">In Process Control</span>
            </a>
        </div>
    </div>
    """, unsafe_allow_html=True)
