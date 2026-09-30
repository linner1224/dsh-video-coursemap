# dsh-video-coursemap

视频课程知识地图 Agent（DeepSeek Harness 插件）。定时观察 Windows 前台浏览器窗口，识别当前播放的视频课程，按学科分类并抓取 B 站简介/字幕/弹幕/评论，生成 Markdown + Mermaid 知识地图与 Word 学习手册。

## 功能

- 每 15 秒轮询一次前台窗口（浏览器/视频客户端），识别视频标题并剥离平台后缀
- LLM 判定「有知识含量」内容（课程/教程/科普/纪录片/测评/知识分享等）并按学科分类
- 按标题反查 B 站视频，抓取简介/字幕/弹幕/评论用于总结（失败自动降级为仅标题）
- 每个学科输出 `学习手册.docx`，根目录输出 `课程总览.docx`（Markdown + Mermaid 思维导图）
- 提供 `course_map_status` / `course_map_scan` 两个 Agent 工具

## 安装

```sh
dsh plugin --profile desktop add github:linner1224/dsh-video-coursemap
```

## 配置

```yaml
- id: dsh-video-coursemap
  name: dsh-video-coursemap
  config:
    enabled: true
    pollIntervalMs: 15000
    stablePolls: 2
    fetchContent: true
    outputDir: null   # null → <用户目录>/课程知识地图
```

## 输出

- `<outputDir>/<学科>/学习手册.docx`：分学科知识手册
- `<outputDir>/课程总览.docx`：全部学科总览
- `<outputDir>/.state/`：内部状态与 Markdown 源

## 限制

- 仅支持 Windows（前台窗口读取依赖 PowerShell + user32.dll）
- 仅在前台窗口是浏览器/视频客户端时生效；后台播放无法检测
- B 站 AI 总结接口需登录，暂不可用；评论匿名接口封顶 3 条

## License

MIT
