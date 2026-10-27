"""Componente HTML de la página Control de obra (dibujo en componentes/control_obra/index.html).
Streamlit exige declarar el componente desde un módulo, no desde el script de la página."""
import os

import streamlit.components.v1 as components

_RUTA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "componentes", "control_obra")
control_obra = components.declare_component("control_obra", path=_RUTA)
