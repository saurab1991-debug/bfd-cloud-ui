
import streamlit as st
from bfd_engine import BFDEngine

st.title("BFD Cloud UI")

batch = st.number_input("Batch size (kg of SM)", value=100.0)

if st.button("Run Calculation"):
    data = {
        "project": {"product_name": "Test Product", "batch_size": batch},
        "components": [
            {"name":"SM","mw":100,"molar_ratio":1,"purity":100,"density":1000,"role":"SM"},
            {"name":"Product","mw":120,"molar_ratio":0.8,"purity":100,"density":1000,"role":"product"}
        ],
        "operations": [
            {"type":"reaction","conversion":100,"selectivity":100}
        ]
    }

    engine = BFDEngine()
    result = engine.calculate(data)

    st.success("Calculation complete")
    st.json(result["yield"])
