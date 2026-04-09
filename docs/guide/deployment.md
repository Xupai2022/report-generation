# 运维部署

本文覆盖当前仓库里能够被代码和部署样例直接证明的服务器运行方式、运行时配置与注意事项。

## 1. 当前仓库支持的部署形态

从代码与 Vite 配置可以直接确认，当前项目的典型部署方式是：

1. 在 `ui-src/` 中构建前端
2. 构建产物写入 `mss_ai_ppt_sample_assets/backend/frontend/`
3. 启动 FastAPI / uvicorn 后端
4. 由后端同时提供：
   - `/api/v1/...` API
   - `/ws/{client_id}` WebSocket
   - `/ui`、`/assets`、`/i18n` 静态页面
   - `/static/previews` 与 `OUTPUTS_URL_PREFIX`

来源：

- 前端输出：`ui-src/vite.config.ts:6`
- 后端挂载：`mss_ai_ppt_sample_assets/backend/app.py:82`

## 2. 服务器上的最小部署步骤

### 构建前端

```bash
cd ui-src
npm install
npm run build
```

### 启动后端

当前仓库与 systemd 样例一致的命令为：

```bash
uvicorn mss_ai_ppt_sample_assets.backend.app:app --host 0.0.0.0 --port 8000
```

这个命令可同时回指到：

- `app.py` 中的本地入口
- `deploy/systemd/mss-ai-ppt.service` 中的样例 `ExecStart`

来源：

- `mss_ai_ppt_sample_assets/backend/app.py:323`
- `deploy/systemd/mss-ai-ppt.service:9`

## 3. 运行时配置

### `.env` 文件位置

当前后端在启动时会：

- 默认从仓库根目录 `.env` 读取配置
- 也支持通过 `MSS_ENV_PATH` 指向其他路径

来源：`mss_ai_ppt_sample_assets/backend/config.py:6`

这意味着服务器部署时，通常至少要确认：

- `.env` 是否存在
- 服务进程是否能读取该文件
- `MSS_ENV_PATH` 是否指向了正确路径

### 路径类配置

当前运行时落盘位置由配置控制，主要包括：

- `MSS_DATA_DIR`
- `MSS_OUTPUTS_DIR`
- `MSS_OUTPUTS_URL_PREFIX`

默认关系：

- `DATA_DIR` -> `backend/data/`
- `OUTPUTS_DIR` -> `backend/outputs/`
- `OUTPUTS_URL_PREFIX` -> `/outputs`

来源：`mss_ai_ppt_sample_assets/backend/config.py:13`

### outputs 下的主要运行目录

当前代码中明确存在以下目录约定：

- `reports/`
- `logs/`
- `previews/`
- `slidespecs/`
- `sessions/`
- `jobs/`
- `rag/`

这意味着部署服务器时，需要关注磁盘空间、权限与备份策略，而不应只盯着 API 可否启动。

来源：`mss_ai_ppt_sample_assets/backend/config.py:18`

## 4. 中转机部署

中转机使用方式：

内网访问https://baolei.sangfor.org/login，用户名为工号，密码大写+小写+数字+十位数，然后点击运维调用mobaxterm

中转机centos7使用Nginx流程：
一、安装 Nginx

yum install -y epel-release
yum install -y nginx

二、启动并设置开机自启

systemctl enable nginx
systemctl start nginx

三、确认服务运行

systemctl status nginx

四、配置反向代理（核心）

编辑配置：

vim /etc/nginx/conf.d/ai_report.conf

写入（标准反代模板）：

server {
listen 8000;
server_name _;

location / {
proxy_pass http://xx.xx.xx.xxx:xxxx;

proxy_set_header Host $host;
proxy_set_header X-Real-IP $remote_addr;
proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
proxy_set_header X-Forwarded-Proto $scheme;

proxy_connect_timeout 60s;
proxy_read_timeout 300s;
}
}

五、检查配置语法

nginx -t

六、重载 Nginx

systemctl reload nginx

七、验证

curl -L http://127.0.0.1:8000

curl http://xx.xx.xx.xxx:xxxx







内网机使用方式：

原网站是https://10.64.0.3/login，现在直接mobaxterm登录即可

装包命令pip install -r mss_ai_ppt_sample_assets/backend/requirements.txt   -i http://mirrors.sangfor.com/nexus/repository/pypi/simple   --trusted-host mirrors.sangfor.com

虚拟环境venv，无需指定python3目录直接source

内网机修改代码后流程：

systemctl daemon-reload

systemctl restart mss-ai-ppt.service

systemctl status mss-ai-ppt.service

内网机防火墙策略：

只允许中转机和我的机器访问

sudo ufw allow ssh (保证不会被锁)

sudo ufw allow from xxx

sudo ufw delete allow from xxx

## 6. 运行时保留与清理策略

当前服务启动时会自动执行：

- session 清理
- preview 临时目录清理
- 旧 preview 清理
- stale lock 清理
- 中断任务恢复标记
- 旧 job 清理

相关保留天数由以下配置控制：

- `PREVIEW_CLEANUP_DAYS`
- `SESSION_RETENTION_DAYS`
- `JOB_RETENTION_DAYS`

这意味着运维上应注意：

- outputs 目录内容不是永久保留
- 重启后 running job 可能会被标记为中断失败
- 清理策略应结合业务对留存时长的要求确认

来源：

- `mss_ai_ppt_sample_assets/backend/config.py:79`
- `mss_ai_ppt_sample_assets/backend/app.py:246`

## 7. RAG 部署注意事项

当前仓库已经支持 RAG 相关运行配置，并且可选在启动时预热。若服务器启用了 RAG，建议额外确认：

- 向量后端配置是否正确
- `RAG_SOURCE_DIR` 是否存在
- 模型文件与磁盘空间是否满足要求
- 是否需要在启动时开启 `RAG_PRELOAD_ON_STARTUP`

来源：

- `mss_ai_ppt_sample_assets/backend/config.py:103`
- `mss_ai_ppt_sample_assets/backend/app.py:248`

## 8. 部署后建议做的最小核对

建议至少核对以下项目：

1. 前端构建产物是否存在于 `backend/frontend/`
2. `/` 是否跳转到 `/ui/index.html`
3. `/api` 是否返回 API 概览
4. `/api/v1/system/health` 是否正常
5. `/api/v1/system/health/detailed` 中 LibreOffice 是否正常
6. 选择一个真实模板与输入后能否走完生成、预览、PPT 下载、PDF 下载

## 9. 环境依赖项的真实性说明

本文中以下内容都属于“需按部署环境校验”的范围：

- 服务器绝对路径
- 服务运行账号
- 反向代理、域名与 HTTPS 方案
- `.env` 的真实部署位置
- `MSS_OUTPUTS_DIR`、`MSS_DATA_DIR` 是否重定向到外挂磁盘
- LibreOffice 安装路径
- RAG 向量后端与模型目录

仓库代码只证明了这些能力与变量的存在，不证明你的服务器已经按这些方式配置完成。
