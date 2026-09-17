"""Application entry point.

macOS 上 Tk 的高分辨率渲染由系统自动完成，不需要 Windows 的
SetProcessDpiAwareness / manifest，因此这里没有平台专有的启动前处理。
若日后要在 Windows 上共用此入口，把 enable_high_dpi() 的 Windows 分支加回即可。
"""

from metadata_editor.ui import launch


if __name__ == "__main__":
    launch()
