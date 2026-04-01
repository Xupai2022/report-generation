FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# ==============================================
# 【关键修复】清空所有源 + 只保留深信服内网源
# ==============================================
RUN rm -rf /etc/apt/sources.list.d/* \
    && rm -rf /etc/apt/sources.list \
    && echo "deb http://mirrors.sangfor.com/nexus/repository/apt-repo/ trixie main" > /etc/apt/sources.list \
    && echo "deb http://mirrors.sangfor.com/nexus/repository/apt-repo/ trixie-updates main" >> /etc/apt/sources.list \
    && echo "deb http://mirrors.sangfor.com/nexus/repository/apt-repo/ trixie-security main" >> /etc/apt/sources.list

# 系统依赖
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    g++ \
    libxml2-dev \
    libxslt1-dev \
    libffi-dev \
    fonts-dejavu-core \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# 安装 LibreOffice
RUN apt-get update && apt-get install -y --no-install-reappends \
    libreoffice \
    libreoffice-impress \
    && rm -rf /var/lib/apt/lists/*

# 确认 soffice
RUN which soffice

# Python 依赖
COPY mss_ai_ppt_sample_assets/backend/requirements.txt /tmp/requirements.txt

RUN pip install --upgrade pip \
    && pip install -r /tmp/requirements.txt

# 复制代码
COPY mss_ai_ppt_sample_assets /app/mss_ai_ppt_sample_assets

EXPOSE 8000

# 启动
CMD ["uvicorn", "mss_ai_ppt_sample_assets.backend.app:app", "--host", "0.0.0.0", "--port", "8000"]