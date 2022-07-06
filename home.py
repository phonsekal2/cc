import streamlit as st
import pandas as pd
from PIL import Image

APP_TITLE = "Dede Saputra App"
st.set_page_config(
    page_title=APP_TITLE,
    page_icon=Image.open("./utils/omic_learn.ico"),
    layout="centered",
    initial_sidebar_state="auto",
)
st.title("Tampak Depan")