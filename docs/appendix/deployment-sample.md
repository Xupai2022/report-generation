# 部署样例说明

本文解释仓库中的 `deploy/systemd/mss-ai-ppt.service`。该文件是**仓库内的部署样例**，不是所有服务器都可直接照抄的最终配置。

## 1. 样例文件位置

- `deploy/systemd/mss-ai-ppt.service`

## 2. 样例中当前可直接确认的内容

```ini
[Service]
Type=simple
WorkingDirectory=/root/report-generation/deploy/systemd
Environment="MSS_ENV_PATH=/opt/report-generation/.env"
ExecStart=/usr/bin/python3 -m uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000
Restart=always
RestartSec=5
LimitNOFILE=65535
```

来源：`deploy/systemd/mss-ai-ppt.service:4`

## 3. 哪些内容可以作为通用参考

### uvicorn 启动命令

以下部分可作为当前后端真实入口的参考：

```ini
ExecStart=/usr/bin/python3 -m uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000
```

这与代码中的本地启动入口一致，说明当前服务确实是以 `mss_ai_ppt_sample_assets.backend.app:app` 作为 uvicorn 应用对象。

来源：

- `deploy/systemd/mss-ai-ppt.service:9`
- `mss_ai_ppt_sample_assets/backend/app.py:323`

### `.env` 可通过 `MSS_ENV_PATH` 指定

样例中的：

```ini
Environment="MSS_ENV_PATH=/opt/report-generation/.env"
```

可以作为“支持外部指定 `.env` 路径”的参考，因为 `config.py` 中确实会读取 `MSS_ENV_PATH`。

来源：

- `deploy/systemd/mss-ai-ppt.service:8`
- `mss_ai_ppt_sample_assets/backend/config.py:7`

## 4. 哪些值属于环境专用，不能直接当成通用事实

### `WorkingDirectory=/root/report-generation/deploy/systemd`

这是样例服务器上的绝对路径，**很可能是环境专用值**。仓库代码并没有要求必须使用这个目录。

来源：`deploy/systemd/mss-ai-ppt.service:7`

### `MSS_ENV_PATH=/opt/report-generation/.env`

这同样是样例环境中的部署路径，不应写成所有服务器都必须使用该位置。

来源：`deploy/systemd/mss-ai-ppt.service:8`

### `/usr/bin/python3`

这取决于服务器上的 Python 安装位置，也不应当作所有环境通用值。

来源：`deploy/systemd/mss-ai-ppt.service:9`

## 5. 使用该样例时建议额外核对什么

如果你要在服务器上参考该文件，至少应重新确认：

- Python 可执行文件路径是否正确
- 服务工作目录是否正确
- `MSS_ENV_PATH` 指向的 `.env` 是否存在
- 服务账号是否有权限读取代码目录、`.env`、outputs 目录
- 前端静态产物是否已经构建到 `backend/frontend/`
- LibreOffice 是否安装并可被服务进程找到

## 6. 文档中的使用原则

因此，知识库中对该样例的引用统一遵守以下边界：

- 可把 `uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000` 作为当前后端入口参考
- 不把 `WorkingDirectory`、`.env` 绝对路径、Python 绝对路径写成通用最终配置
- 所有绝对路径类值都需要明确标注“需按部署环境校验”
