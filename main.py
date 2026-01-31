import streamlit as st
import pandas as pd
from pathlib import Path

st.write('Hello World')

base_path = Path("data_bases/transit_data.XLS")
df = pd.read_excel(base_path)

st.write(df)