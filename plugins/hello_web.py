# plugins/hello_web.py — Web対応プラグイン例
from app_shell import PluginBase

class Plugin(PluginBase):
    name = "Hello Web"
    icon = "🌐"

    def web_mount(self, st):
        st.write("Hello from web plugin!")
        txt = st.text_input("Type something", "Streamlit is easy")
        if st.button("Show"):
            st.success(f"You typed: {txt}")
