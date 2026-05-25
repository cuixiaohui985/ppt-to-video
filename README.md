# PPT to Video

将 PPT/PPTX 文件转换为带背景音乐的 MP4 视频，支持智能时长、翻页动画和打字机文字渐入效果。

## 三大核心功能

| 功能 | 说明 |
|---|---|
| **智能时长** | 封面/过渡页 1.5s，封底 4.0s，正文页按 180字/分（3字/秒）计算，范围 6~8s |
| **翻页动画** | 向上翻页（slideup）、对角线翻页（diagtl）、淡入淡出（fade）、无动画（none） |
| **打字机效果** | 正文页文字逐字渐入，PIL 渲染透明文字图 + ffmpeg crop 渐进裁剪 |

## 快速开始

```bash
# 基础用法（上翻动画，无打字机效果）
python scripts/ppt_to_video.py presentation.pptx bg_music.mp3 --transition slideup

# 启用打字机效果 + 对角线翻页
python scripts/ppt_to_video.py presentation.pptx bg_music.mp3 --typewriter --transition diagtl

# 自定义参数
python scripts/ppt_to_video.py presentation.pptx bg_music.mp3 \
  --cover-dur 2.0 \
  --content-cps 2.5 \
  --content-min 5.0 \
  --content-max 10.0
```

## 依赖安装

```bash
pip install python-pptx pillow pywin32 pdf2image
```

系统需安装 **ffmpeg**（用于视频/音频合成）：

| 系统 | 安装方式 |
|---|---|
| Windows (scoop) | `scoop install ffmpeg` |
| macOS (Homebrew) | `brew install ffmpeg` |
| Linux | `sudo apt install ffmpeg` |

## PPT 渲染方案

脚本自动按优先级选择最佳渲染方案：

| 优先级 | 方案 | 质量 | 依赖 |
|---|---|---|---|
| 1 | COM 自动化（300DPI） | ⭐⭐⭐⭐⭐ | PowerPoint + pywin32 |
| 2 | LibreOffice PDF→PNG（300DPI） | ⭐⭐⭐⭐ | LibreOffice + pdf2image |
| 3 | LibreOffice 直接 PNG | ⭐⭐⭐ | LibreOffice |

## 全部参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `ppt` | （必填） | PPT/PPTX 文件路径 |
| `audio` | （必填） | 背景音乐文件路径 |
| `--output` | 同目录 .mp4 | 输出视频路径 |
| `--transition` | `slideup` | 翻页动画：slideup/diagtl/fade/none |
| `--transition-dur` | `0.5` | 翻页动画时长（秒） |
| `--typewriter` | 不启用 | 正文页打字机效果 |
| `--cover-dur` | `1.5` | 封面/过渡页时长（秒） |
| `--back-cover-dur` | `4.0` | 封底时长（秒） |
| `--content-cps` | `3.0` | 正文页语速（字/秒） |
| `--content-min` | `6.0` | 正文页最短时长（秒） |
| `--content-max` | `8.0` | 正文页最长时长（秒） |
| `--fps` | `25` | 输出视频帧率 |
| `--keep-temp` | 不保留 | 保留临时文件（调试用） |

## 作为 WorkBuddy Skill 使用

本仓库同时是一个 [WorkBuddy](https://www.codebuddy.cn/) Skill。

触发词：`PPT转视频`、`PPT生成视频`、`PPT加音乐`、`PPT to video`

安装方式：将整个目录放入 `~/.workbuddy/skills/ppt-to-video/`

详细文档见 [SKILL.md](SKILL.md)。

## 许可

MIT License
