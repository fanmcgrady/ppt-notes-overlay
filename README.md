# 答辩专用插件

Windows 下的 PowerPoint 答辩/录屏辅助工具：放映时置顶显示当前页备注，并提供排练计时与备注导出。

> 不是 Office 官方加载项，而是通过 COM 连接桌面版 PowerPoint 的独立小工具。

## 功能

- 放映 / 录制时浮动显示**当前页备注**（事件驱动 + 高频兜底，尽量实时）
- **本页计时 / 总计时**（开始放映自动计时，结束自动暂停）
- **置顶**开关、字号放大缩小（`A+` / `A−` 或 `Ctrl+滚轮`）
- 窗口可拖拽改大小
- **逐页导出备注为 Word（.docx）**
- 翻页使用 PowerPoint 原生快捷键（`→` / `←` / 空格等）

## 环境要求

- Windows 10 / 11
- 桌面版 Microsoft PowerPoint
- Python 3.10+（建议 3.11+）

## 快速开始

```bat
pip install -r requirements.txt
python notes_overlay.py
```

或双击 `scripts\run.bat`。

使用步骤：

1. 先打开带备注的 PPT
2. 运行本工具
3. 在 PowerPoint 中开始放映 / 录制
4. 浮动窗会同步显示当前页备注与计时

录制时建议只录制/共享幻灯片窗口，避免把备注浮窗一起录进去。

## 打包成 exe

```bat
scripts\build_exe.bat
```

产物：`dist\答辩专用插件.exe`

## 目录结构

```text
ppt-notes-overlay/
├── notes_overlay.py      # 主程序
├── assets/
│   ├── icon.ico          # 应用图标
│   └── icon.png
├── scripts/
│   ├── run.bat           # 开发运行
│   └── build_exe.bat     # 打包 exe
├── requirements.txt
├── requirements-dev.txt
├── LICENSE
└── README.md
```

## 说明与免责

- `assets/icon.*` 默认图标如使用校徽/商标，请在公开发布前确认授权，或自行替换。
- 本工具按 MIT 协议开源，不提供任何担保；答辩/录屏请自行做好备份与演练。

## License

MIT
