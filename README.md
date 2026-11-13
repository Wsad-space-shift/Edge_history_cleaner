# Edge 站点历史清理器

> [!waring]
> vibe警告

按“名单”删除 Microsoft Edge 浏览记录里 **指定站点** 的痕迹，其它站点的记录保持原样。

和“清空全部历史”不同：它只挑你点名的域名，适合处理一两个不想留痕的站点，
而不必牺牲整段浏览记录。

```
双击  启动清理器.bat      → 图形界面
```

## 它清什么

| 位置 | 内容 |
|---|---|
| `History` | 网址、访问、搜索词、访问分段、来源(referrer)字段、下载记录、Edge 的互动/内容标注 |
| `Top Sites` | 新标签页和地址栏的“常用站点” |
| `Shortcuts` | 地址栏下拉里的历史提示 |
| `Network Action Predictor` | 地址栏输入预测、资源预取记录 |
| `Favicons` | 这些页面的图标映射与图标本体 |

清理采用 `secure_delete` + `VACUUM`，删除的数据块会被重写掉，不会以裸数据留在文件里。

## 名单怎么写

每行一条，空行和 `#` 开头的行忽略：

```text
# 社交
x.com            # 含 www.x.com、api.x.com 等子域
twitter.com
t.co

# 精确到某个子域
mail.example.com

# 整条网址也行，只取主机名
https://news.example.com/path?q=1
```

- 默认“域名 + 全部子域”；取消勾选“同时匹配其子域”则只认填的那个主机名。
- `*.example.com` 写法可强制包含子域。
- 无法识别的行会在日志里提示，不会让整次清理失败。

## 使用流程

1. 填名单 → 选配置文件（`Default`、`Profile 1`…，可 Ctrl 多选）。
2. 点 **预览匹配量**：只复制数据库来统计，不改任何文件，可以反复看。
3. 点 **开始清理**：确认后自动关闭 Edge → 备份 → 逐个库清理 → `VACUUM` → 重开 Edge。
4. 结果窗口会给出每个表的删除条数、数据库体积变化和完整性检查。

出问题时用 **从备份恢复…** 选中那次备份即可回到清理前的状态（同样需要关闭 Edge）。

## 命令行

```powershell
# 看看会清掉多少（不改数据）
python edge_history_cleaner.py --site x.com --site t.co --dry-run

# 真的清
python edge_history_cleaner.py --site x.com --site t.co

# 名单文件 + 指定配置文件 + 只清 History
python edge_history_cleaner.py --list-file 名单.txt --profile Default --no-stores

# 从某次备份恢复
python edge_history_cleaner.py --restore "C:\Users\你\Documents\Edge历史清理备份\20261113-065954"
```

`python edge_history_cleaner.py --help` 有全部参数。

## 备份

默认写到 `文档\Edge历史清理备份\<时间戳>\<配置文件>\`，包含本次涉及的 5 个数据库、
`Preferences` 和 `manifest.json`（记录目标站点、时间、配置文件）。
备份目录可以在界面上改。备份里有未清理的原始记录，确认没问题后自行删除。

## 已知边界

- **最近关闭的标签页 / 会话恢复**：数据在 `Sessions\Tabs_*`、`Session_*` 里，不在本次范围内。
  清掉会连带丢失“恢复上次会话”，所以留给你自己决定。
- **站点参与度**：`Preferences` 里还有 `https://x.com:443` 之类的 engagement 分数，
  属于统计信息，不影响历史列表，本工具不动它。
- **账号同步**：若 Edge 开着历史同步，云端可能仍有一份；实测本地清完重启后没有回填。
  如果发现记录重新出现，去 `edge://settings/profiles/sync` 关掉历史同步再清一次。
- 清理前后 Edge 会被关闭/重启，未保存的网页表单内容可能丢失。

## 技术说明

- 纯 Python 3（标准库 `sqlite3` + `tkinter`），无第三方依赖，不需要管理员权限。
- 直接识别 Edge 的真实表结构（`urls`/`visits`/`visited_links`/`edge_visits`…），
  对缺少表或字段的 Edge 版本会跳过对应步骤而不是报错。
- 其它数据库的删除按行值解析主机名，避免 `xx.com` 误伤 `proxyxx.com`。
- 预览在临时副本上进行，清理前会 `tasklist` 检测并优雅关闭 Edge，必要时才强制结束。
