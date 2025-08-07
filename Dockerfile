# 使用多阶段构建减少最终镜像大小和安全风险
FROM node:18-alpine@sha256:c698ffe060d198dcc6647be78ea1683363f12d5d507dc5ec9855f1c55966ffdd AS builder

WORKDIR /app

# 只复制必要的文件
COPY package*.json pnpm-lock.yaml ./
RUN npm install -g pnpm@8.15.1 --registry=https://registry.npmjs.org/
RUN pnpm install --frozen-lockfile --prod --registry=https://registry.npmjs.org/

# 第二阶段：运行时镜像
FROM node:18-alpine@sha256:c698ffe060d198dcc6647be78ea1683363f12d5d507dc5ec9855f1c55966ffdd

WORKDIR /app

# 安装运行时依赖
RUN apk add --no-cache \
    chromium=119.0.6045.159-r0 \
    nss=3.94-r0 \
    && rm -rf /var/cache/apk/*

# 从构建阶段复制依赖
COPY --from=builder /app/node_modules ./node_modules

# 只复制必要的运行文件，避免敏感目录
COPY --chown=nodejs:nodejs dist/ ./dist/
# 避免复制整个src目录，只复制必要的文件
# COPY --chown=nodejs:nodejs src/ ./src/
# 使用条件复制，避免敏感文件
RUN mkdir -p ./resources && chown -R nodejs:nodejs ./resources

# 创建非 root 用户并设置权限
RUN addgroup -g 1001 -S nodejs && \
    adduser -S docops -u 1001 && \
    chown -R docops:nodejs /app

USER docops

HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD node dist/index.cjs --health-check || exit 1

CMD ["node", "dist/index.cjs"]