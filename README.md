# dsh-video-coursemap · 视频课程知识地图 Agent 2.0

一个纯 DSH 插件：**实时检测网页端 B 站视频课程 → 按关键词自动学科分类 → 抓取视频内容/弹幕/评论做知识总结 → 输出 Word 知识地图与学习手册**。无需任何运行时依赖，无需修改 DSH 核心。

## 安装

### 方式一：Release 预构建包（推荐，无需 git、无需 npm）

```bash
dsh plugin add "https://github.com/linner1224/dsh-video-coursemap/releases/download/v2.0.0/dsh-video-coursemap-2.0.0.tgz"
```

直接从 GitHub Release 下载打包好的产物，全程不调用 `git`，也不需要 npm 账号或构建工具。这是**兼容性最好**的方式，任何装了 DSH 的机器都能成功。

### 方式二：插件市场

在市场里搜索 `dsh-video-coursemap` 一键安装。市场能否免 git 安装，取决于索引里这条记录是否已带上 `tarball` 映射 —— 详见下方《装不上怎么排查》。

### 方式三：GitHub 地址（需要本机已装 git）

```bash
dsh plugin add github:linner1224/dsh-video-coursemap
```

此方式要求本机**已安装 git 且在 PATH 中**，并且能访问 github.com。缺任意一项都会失败。

## 装不上怎么排查

| 现象 | 原因 | 解决 |
| --- | --- | --- |
| `ERR_PNPM_GIT_RESOLVE_FAILED`、`git ls-remote failed: 'git' 不是内部或外部命令` | 本机没装 git，或 git 不在 PATH | 安装 Git for Windows 并重启 DSH，或改用**方式一** |
| `ETIMEDOUT`、连接 github.com 超时 | 网络无法访问 GitHub | 给 git/pnpm 配代理，或改用**方式一** |
| 市场里点安装失败，但手动 `dsh plugin add` 能装 | 市场索引尚未收录本插件的 `tarball` 映射，退化成了裸 `github:` 地址（必须走 git） | 改用**方式一**，或等索引刷新 |
| 装完 `course_map_*` 工具不出现 | 插件未被加载 | 重启 DSH（DSH 对插件模块做缓存，不会热重载） |
| 能装、能跑，但生成 Word 失败 | 缺少 Python 环境（`scripts/md2docx.py` 依赖本机 python） | 安装 Python，或先关闭 Word 输出 |

## 维护者：发版与「市场免 git 安装」

市场的安装目标由 `installTargetFor` 决定，优先级为：

1. `npm` 包（发布到 npm registry）
2. GitHub Release 上的 `.tgz` 资产（`tarball` 映射）
3. 兜底：`github:owner/repo` ← **要求用户本机有 git**

本插件未发布到 npm，因此**必须挂 Release 资产**，市场才会生成免 git 的安装命令。发版步骤：

1. 打包：`pnpm pack --pack-destination <输出目录>`，得到 `dsh-video-coursemap-<版本>.tgz`
2. 在 GitHub 创建 Release，**tag 与 `package.json` 的 version 保持一致**（如 `v2.0.0`），把 `.tgz` 作为二进制资产上传
3. 等待聚合索引刷新（`awesome-dsh-plugin.com`，其 `build-site` 工作流每天 **10:23（北京时间）** 运行）
4. 验证映射是否生效：

```powershell
$j = Invoke-RestMethod "https://awesome-dsh-plugin.com/plugins.json"
($j.plugins | Where-Object { $_.name -eq "dsh-video-coursemap" }) |
  Select-Object npm, version, tarball, install
```

`tarball` 字段返回 Release 的 `.tgz` 地址即表示成功，此时市场按钮会走免 git 路径。

**注意**：`package.json` 里的 `files` 字段决定打包内容，增删文件时需同步更新；`private` 字段会阻止发布到 npm（不影响 Release 资产）。

## 2.0 三大功能

### 功能一：实时检测 & 静默去重
- 定时（默认每 15s）读取 Windows **前台窗口标题**，识别浏览器/视频客户端中的视频。
- **新视频** → 连续 `stablePolls` 次（默认 2 次）稳定确认后自动生成知识地图；
- **已处理过的视频** → 保持**静默**，不重复生成（记录到 `seen` 去重表）；
- 非知识内容（纯音乐、搞笑段子、无解说实况、直播、新闻、普通网页等）→ 静默忽略。

### 功能二：学科分类
- 由 LLM 根据视频**标题 + 关键词**判断所属学科（中文 ≤8 字）；
- **旧学科** → 追加更新到已有的知识地图 / 学习手册；
- **新学科** → 自动新建该学科的学习手册。

### 功能三：内容知识总结 & Word 输出
- 按标题反查 B 站视频，抓取：视频简介、标签、字幕（如有）、**弹幕**（≤150 条）、**热评**（≤20 条）；
- 基于真实内容 + 弹幕 + 评论提炼**核心知识点**（8~20 条、2~3 层）与**思维导图**（Mermaid mindmap）；
- 输出为 **Word（.docx）**：每个学科一本《学习手册.docx》+ 全局《课程总览.docx》；
- Mermaid 思维导图在 Word 里渲染为**原生分级大纲**（根加粗 + 分级项目符号），而非代码文本。

## 输出目录结构

```
<输出目录>（默认 C:\Users\<你>\课程知识地图，可用 outputDir 配置）
├── 课程总览.docx            # 全局知识地图（学科总览 + 学科列表）
├── <学科A>/
│   └── 学习手册.docx        # 该学科所有视频的知识总结
├── <学科B>/
│   └── 学习手册.docx
└── .state/                  # 内部状态（state.json、manuals/*.md 源、poll.log 调试日志）
```

## 配置项

| 配置 | 默认 | 说明 |
| --- | --- | --- |
| `enabled` | `true` | 总开关 |
| `pollIntervalMs` | `15000` | 轮询间隔（毫秒） |
| `stablePolls` | `2` | 连续 N 次观察到同一新标题才处理（防标签切换抖动） |
| `outputDir` | `null` | 输出目录，`null` → `<用户目录>/课程知识地图` |
| `provider` / `model` | `null` | LLM 提供方/模型，`null` → 跟随 DSH 默认模型 |
| `maxTokens` | `4096` | 生成知识地图的最大 token |
| `titleMinLen` | `4` | 标题最小长度（过滤空/短标题） |
| `fetchContent` | `true` | 是否抓取 B 站简介/字幕/弹幕/评论（失败自动降级为仅标题总结） |

## 提供的工具

| 工具 | 作用 |
| --- | --- |
| `course_map_status` | 查看 Agent 状态：是否启用、输出目录、已分类学科与视频数、最近观察标题 |
| `course_map_scan` | 立即执行一次前台窗口观察与知识地图生成（可传 `title` 直接指定视频标题，便于测试） |

## 技术要点 / 坑

- **改插件代码后必须重启 DSH** 才生效（DSH 对插件模块做缓存，remove+reinstall 不会热重载）。
- B 站 API 需带浏览器 User-Agent + Referer（否则搜索返回 `-412`）；通过 `/x/frontend/finger/spi` 取 `buvid3` 提升稳定性。
- 评论匿名接口封顶 3 条（`sort=2` 热评、`ps` 上限 20）；AI 总结接口 `/x/web-interface/view/conclusion/get` 返回 `-403` 需登录，不可用。
- 多数 B 站视频无机器字幕（`subtitle` 为空），字幕是尽力抓取、失败静默降级。
- 前台窗口必须是浏览器/视频客户端才能识别；后台播放（窗口不在前台）属预期限制。

## 目录

```
dsh-video-coursemap/
├── package.json         # dsh.bundle.patch → cordis.patch.yml
├── cordis.patch.yml     # 把插件行 insert 进 profile 插件清单
├── lib/index.js         # Agent 主逻辑（轮询/分类/抓取/生成/Word 转换/工具）
└── scripts/md2docx.py   # Markdown → Word（含 mermaid mindmap → 原生分级大纲）
```
