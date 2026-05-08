# ai-dev-share-frontend-slides

一套使用零依赖 HTML slides 制作的内部 AI 分享演示稿。

主题聚焦：

- Codex
- Claude Code
- Skills
- MCP
- qclaw
- WorkBuddy

目标不是做产品说明书，而是分享真实使用经验、技巧、踩坑和协作方式。

## 本地打开

直接用浏览器打开 [index.html](/Users/wsm/Project/ai-dev-share-frontend-slides/index.html) 即可。

如果你想走本地静态服务：

```bash
cd /Users/wsm/Project/ai-dev-share-frontend-slides
python3 -m http.server 4173
```

然后访问 `http://localhost:4173`。

## 当前结构

- `index.html`：完整 slides 页面
- `assets/`：截图素材

## 下一步

- 增加 qclaw / WorkBuddy 的真实案例页
- 增加 Skills / MCP 的流程图页
- 处理截图脱敏与压缩
- 发布到 GitHub Pages 或 Vercel
