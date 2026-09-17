#!/bin/sh
# 把源码打包成可双击运行的 FileMetadataEditor.app
# 用法：sh build_app.sh
set -eu

HERE=$(cd "$(dirname "$0")" && pwd)
APP="$HERE/FileMetadataEditor.app"
MACOS_DIR="$APP/Contents/MacOS"
RESOURCES="$APP/Contents/Resources"

rm -rf "$APP"
mkdir -p "$MACOS_DIR" "$RESOURCES"

cp -R "$HERE/metadata_editor" "$RESOURCES/metadata_editor"
cp "$HERE/app.py" "$RESOURCES/app.py"
find "$RESOURCES" -name '__pycache__' -type d -prune -exec rm -rf {} +
find "$RESOURCES" -name '*.pyc' -delete

cat > "$APP/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>文件元数据编辑</string>
    <key>CFBundleDisplayName</key>
    <string>文件元数据编辑</string>
    <key>CFBundleIdentifier</key>
    <string>local.filemetadataeditor.macos</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>NSHumanReadableCopyright</key>
    <string>本地运行，不联网。</string>
    <key>CFBundleDocumentTypes</key>
    <array>
        <dict>
            <key>CFBundleTypeName</key>
            <string>Office 文档</string>
            <key>CFBundleTypeRole</key>
            <string>Editor</string>
            <key>LSItemContentTypes</key>
            <array>
                <string>org.openxmlformats.wordprocessingml.document</string>
                <string>org.openxmlformats.spreadsheetml.sheet</string>
                <string>org.openxmlformats.presentationml.presentation</string>
                <string>com.adobe.pdf</string>
            </array>
        </dict>
    </array>
</dict>
</plist>
PLIST

# 启动器：从 .app 双击时工作目录不是脚本目录，必须自己定位 Resources 并挑一个带 Tk 的 python3。
cat > "$MACOS_DIR/launcher" <<'LAUNCHER'
#!/bin/sh
DIR=$(cd "$(dirname "$0")/../Resources" && pwd)
cd "$DIR" || exit 1

for CANDIDATE in \
    /usr/local/bin/python3 /opt/homebrew/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/Current/bin/python3 python3
do
    command -v "$CANDIDATE" >/dev/null 2>&1 || continue
    "$CANDIDATE" -c 'import tkinter' >/dev/null 2>&1 || continue
    exec "$CANDIDATE" app.py
done

/usr/bin/osascript -e 'display alert "无法启动" message "未找到带 Tkinter 的 python3。请先安装官方 Python（python.org）或 Homebrew 版 python-tk。"' >/dev/null 2>&1
exit 1
LAUNCHER
chmod +x "$MACOS_DIR/launcher"

echo "已生成：$APP"
echo "双击即可运行；如提示无法打开，执行： xattr -dr com.apple.quarantine \"$APP\""
