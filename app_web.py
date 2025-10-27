# app_web.py  — Streamlit 版エントリ
import importlib.util
from pathlib import Path
import streamlit as st

PLUGINS_DIR = Path(__file__).parent / "plugins"

st.set_page_config(page_title="Plugin Shell (Web)", layout="wide")
st.title("🔌 Plugin Shell (Web)")
st.caption("Streamlit 版。plugins/ に web_mount(st) を持つプラグインだけを表示します。")

def load_web_plugins():
    plugins = []
    PLUGINS_DIR.mkdir(exist_ok=True)
    for py in sorted(PLUGINS_DIR.glob("*.py")):
        try:
            spec = importlib.util.spec_from_file_location(py.stem, py)
            if not spec or not spec.loader:
                continue
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)  # type: ignore
            if hasattr(module, "Plugin"):
                from app_shell import PluginBase  # ルートにある前提
                Plugin = getattr(module, "Plugin")
                if issubclass(Plugin, PluginBase):
                    inst = Plugin(shell_context={"base_dir": str(Path(__file__).parent)})
                    if hasattr(inst, "web_mount"):
                        plugins.append(inst)
        except Exception as e:
            st.warning(f"Failed to load {py.name}: {e}")
    return plugins

plugins = load_web_plugins()
names = [getattr(p, "name", "Unnamed") for p in plugins]

left, right = st.columns([1, 3], gap="large")
with left:
    if not plugins:
        st.info("plugins/ に web_mount() を実装したプラグインを置くとここに出ます。")
    idx = st.radio("プラグイン", options=list(range(len(plugins))), format_func=lambda i: names[i]) if plugins else None

with right:
    if idx is not None:
        plg = plugins[idx]
        st.subheader(f"{getattr(plg, 'icon', '🔹')} {plg.name}")
        plg.web_mount(st)
