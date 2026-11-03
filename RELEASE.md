# 发版操作规范（维护者）

> 目标：每次发版都保证「插件市场能免 git 安装」。本文档是唯一权威流程，发版前逐项打勾。

## 0. 先理解：市场的安装目标是怎么定的

插件市场安装时不会直接拿仓库地址，而是由 `installTargetFor` 按优先级推导一个安装目标：

| 优先级 | 条件 | 生成的安装命令 | 用户端要求 |
| --- | --- | --- | --- |
| 1 | 已发布到 npm（索引 `npm` 字段非空） | `dsh plugin add <包名>` | 能访问 npm registry |
| 2 | Release 挂了 `.tgz` 资产（索引 `tarball` 字段非空） | `dsh plugin add "<release tgz URL>"` | 能访问 github.com |
| 3 | 兜底 | `dsh plugin add github:linner1224/dsh-video-coursemap` | **本机必须装 git** |

**结论：本插件不发布到 npm，所以每次发版必须挂 Release `.tgz` 资产。** 否则落到第 3 档，没装 git 的用户 100% 失败，报错为：

```
ERR_PNPM_GIT_RESOLVE_FAILED  git ls-remote failed: 'git' 不是内部或外部命令
```

## 1. 发版前检查清单

- [ ] `package.json` 的 `version` 已递增（如 `2.0.1`）
- [ ] 计划用的 tag 与 version 严格对应（`2.0.1` ↔ `v2.0.1`）
- [ ] `package.json` 中没有 `"private": true`
- [ ] `files` 字段覆盖全部运行时文件（`lib`、`cordis.patch.yml`、`README.md`、`scripts`）
- [ ] `lib/`、`cordis.patch.yml`、`scripts/` 的改动已提交到 `main`
- [ ] **`README.md` 里「方式一」的下载 URL 已改成新版本号**（该 URL 写死了版本，最容易漏）
- [ ] `LICENSE` 等仓库独有文件未被误删

## 2. 打包

用 DSH 自带的 pnpm 打包，产物名必须与版本一致：

```powershell
$node  = "C:\Users\ASUS\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\node\bin\node.exe"
$pnpm  = "C:\Users\ASUS\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\pnpm\bin\pnpm.mjs"
$dest  = "C:\Users\ASUS\coursemap-dist"

& $node $pnpm --dir "E:\dsh\dsh-video-coursemap" pack --pack-destination $dest
```

预期输出 `dsh-video-coursemap-<版本>.tgz`，并打印 Tarball Contents —— **核对清单里没有多余文件、也没有漏掉 `lib/index.js`**。

## 3. 创建 Release

网页方式（推荐，不需要本地 git）：

1. 打开 `https://github.com/linner1224/dsh-video-coursemap/releases/new`
2. 「标签」填入 `v<版本>`（如 `v2.0.1`），回车创建
3. 「发行标题」同 tag
4. 「发行说明」写一句话即可
5. 「发行厂牌」选「没有」
6. **把 `.tgz` 拖进最下方「附加二进制文件」区域**，等文件名出现
7. 点「发布发行」

命令行方式（需 `gh auth login`）：

```powershell
$gh = "C:\Users\ASUS\AppData\Local\Programs\GitHub CLI Portable\bin\gh.exe"
& $gh release create "v2.0.1" "C:\Users\ASUS\coursemap-dist\dsh-video-coursemap-2.0.1.tgz" `
    --repo linner1224/dsh-video-coursemap --title "v2.0.1" --notes "预构建分发包"
```

## 4. 发布后校验（三项全过才算发布完成）

```powershell
$v   = "2.0.1"
$url = "https://github.com/linner1224/dsh-video-coursemap/releases/download/v$v/dsh-video-coursemap-$v.tgz"

# ① 资产可下载
$r = Invoke-WebRequest $url -Method Head -UseBasicParsing
"HTTP $($r.StatusCode)  size=$($r.Headers['Content-Length'])"

# ② 哈希与本地一致
(Get-FileHash "C:\Users\ASUS\coursemap-dist\dsh-video-coursemap-$v.tgz" -Algorithm SHA256).Hash

# ③ 免 git 独立安装（最关键）
$t = "C:\Users\ASUS\cm-verify"
Remove-Item $t -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory $t -Force | Out-Null
& $node $pnpm --dir $t add $url
Test-Path "$t\node_modules\dsh-video-coursemap\lib\index.js"   # 必须 True
Remove-Item $t -Recurse -Force
```

## 5. 等待索引刷新并验证

聚合站 `awesome-dsh-plugin/awesome-dsh-plugin` 的定时任务（北京时间）：

| 工作流 | 时间 | 说明 |
| --- | --- | --- |
| `build-site` | 每天 **10:23** | 重建 `plugins.json`，主刷新点 |
| `regate` | 每 6 小时（02:19 / 08:19 / 14:19 / 20:19） | 重校验条目事实 |
| `held-rescan` | 每天 14:19 | 重扫待定条目 |

**若在 10:23 之后才发布，就要等次日，最长约 24 小时。** 可用 `workflow_dispatch` 手动触发（需仓库权限），或提 issue 请维护者刷新。

验证映射是否生效：

```powershell
$j = Invoke-RestMethod "https://awesome-dsh-plugin.com/plugins.json"
($j.plugins | Where-Object { $_.name -eq "dsh-video-coursemap" }) |
  Select-Object npm, version, tarball, install
```

判定标准：

- `tarball` = `https://github.com/linner1224/dsh-video-coursemap/releases/download/v<版本>/dsh-video-coursemap-<版本>.tgz`
- `install` 不再是 `github:...` 形式

## 6. 故障处理

| 现象 | 原因 | 处理 |
| --- | --- | --- |
| Release 发布后索引仍为 `tarball: null` | 未到刷新时间，或爬虫未识别资产 | 等下一个 10:23；确认资产挂在**本仓库** Release 下 |
| 索引 `version` 与 tag 不一致 | `main` 上 `package.json` 的 version 未更新 | 提交 package.json 到 main 后等刷新 |
| 用户安装报 `ERR_PNPM_GIT_RESOLVE_FAILED` | 索引尚未填 `tarball`，退化成 `github:` | 让用户改用 README「方式一」的直链 |
| 用户安装报下载失败/超时 | 用户网络访问不了 GitHub | 配代理；根治需发布到 npm 或镜像源 |
| 需要撤回某次发版 | —— | Release 页面删除该资产或将 Release 标记为「预发行」，市场随之回退 |

## 7. 完成标准（Definition of Done）

一次发版只有同时满足这三条才算结束：

1. Release 资产可下载，SHA256 与本地构建一致
2. 在**无 git 的干净环境**中可用 Release URL 安装成功
3. 索引中 `tarball` 字段已指向该资产

第 3 条依赖聚合站刷新，通常延迟到次日；未达成前，README「方式一」的直链是对用户唯一可靠的安装方式。