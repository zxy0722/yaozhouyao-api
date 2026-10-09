# 耀州窑青釉刻花瓶知识 API —— 云端部署指南（零成本验证）

本指南让 API 在 **5~10 分钟内**上线到公网，全程 0 元。
原理：代码已打包成 Docker 镜像（`Dockerfile`），通过 Render 免费容器托管直接跑起来。

---

## 一、把代码推送到 GitHub（约 5 分钟）

> 仓库已在本机初始化并完成首次提交，你只需"建远程仓库 + 推送"。

1. 打开 https://github.com/new 登录（无账号先注册，免费）。
2. 仓库名填：`yaozhou-kiln-api`，选择 **Public**，不要勾选任何初始化选项，点击 **Create repository**。
3. 创建成功后页面会显示一段命令，复制并执行（在你项目目录的终端里）：

```bash
git remote add origin https://github.com/<你的用户名>/yaowl-api.git
git branch -M main
git push -u origin main
```

> 也可以不用命令行：直接在这个仓库页面点 **uploading an existing file**，把
> `main.py、schemas.py、requirements.txt、data/dataset.json、guide.html、Dockerfile、render.yaml、.gitignore`
> 这些文件拖进去上传，效果相同。

## 二、Render 一键部署（约 3 分钟）

1. 打开 https://dashboard.render.com 注册（推荐用 GitHub 账号直接登录）。
2. 点击 **New → Blueprint**，选择你刚创建的 `yaowl-api` 仓库。
3. Render 会自动读取 `render.yaml`，创建名为 `yao-kiln-api` 的免费 Web 服务。
4. 点击 **Apply**，等待 3-5 分钟构建完成，状态变绿 `Live`。

> 如果 Blueprint 没识别到，可改用 **New → Web Service → 选仓库**，Runtime 选 **Docker**，其余默认。

---

## 三、验证接口（部署完成后）

服务上线后会分配一个公网域名，形如：
`https://yao-kiln-api.onrender.com`

逐个验证（把域名替换进浏览器）：

| 接口 | 用途 | 期望结果 |
|---|---|---|
| `GET /health` | 健康检查 | `{"status":"ok"}` |
| `GET /api/v1/artifacts` | 文物列表 | code=0，返回文物摘要 |
| `GET /api/v1/artifacts/yz-001/records` | 记录检索 | 分页数据 |
| `GET /api/v1/stats` | 数据统计 | 各类别计数 |
| `GET /api/v1/search?q=釉` | 全文搜索 | 命中结果 |
| `GET /api/v1/narrate?q=纹饰` | AI 讲解词 | 文字讲解 |
| `POST /api/v1/tts` | 语音合成 | 返回 mp3 |
| `GET /guide` | 智能导览页 | 页面可打开 |
| `GET /docs` | Swagger 文档 | 可视化调试页面 |

也可以在浏览器直接打开 `https://<你的域名>/docs`，在 Swagger 界面点接口试跑。

---

## 四、注意事项

- **免费额度**：免费实例每月 750 小时（一个实例约 31 天），完全够验证使用。
- **休眠唤醒**：免费实例闲置 15 分钟会进入休眠，再次访问需等待 30–60 秒自动唤醒，属正常现象。
- **访问速度**：Render 免费域名在境内访问速度尚可，适合验证；如需长期稳定使用，建议后续用同一套 `Dockerfile` 部署到天翼云/阿里云国内服务器（几十元/年），响应更快。
- **数据安全**：`data/dataset.json` 数据来源于故宫博物院等公开资料，已在代码中注明出处，符合公开资料使用规范。
- **可选增强**：如需 AI 讲解的大模型润色功能，在 Render 服务环境变量中配置 `LLM_BASE_URL`、`LLM_API_KEY`、`LLM_MODEL` 即可（不配则自动用模板兜底，不影响使用）。

---

## 五、清理与下线（不需要时）

在 Render 控制台选择服务 → **Delete**，即可删掉云端实例，不会产生任何费用。

> AI生成
