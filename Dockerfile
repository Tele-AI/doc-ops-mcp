# Use official Node.js runtime as base image with fixed hash for security
FROM node:18-alpine@sha256:c698ffe060d198dcc6647be78ea1683363f12d5d507dc5ec9855f1c55966ffdd

# Set working directory
WORKDIR /app

# Install system dependencies for headless Chrome with version pinning
RUN apk add --no-cache \
    chromium=119.0.6045.159-r0 \
    nss=3.94-r0 \
    freetype=2.13.0-r5 \
    freetype-dev=2.13.0-r5 \
    harfbuzz=8.2.2-r0 \
    ca-certificates=20230506-r0 \
    ttf-liberation=2.1.5-r1 \
    && rm -rf /var/cache/apk/*

# Set environment variables for Playwright
ENV PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1
ENV PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH=/usr/bin/chromium-browser
ENV PLAYWRIGHT_SKIP_VALIDATE_HOST_REQUIREMENTS=true

# Create app directory structure
RUN mkdir -p /app/dist /app/documents /app/resources /app/temp

# Copy package files
COPY package*.json ./
COPY pnpm-lock.yaml ./

# Install pnpm with fixed version to prevent dependency confusion
RUN npm install -g pnpm@8.15.1 --registry=https://registry.npmjs.org/

# Install dependencies with explicit registry to prevent dependency confusion
RUN pnpm install --frozen-lockfile --prod --registry=https://registry.npmjs.org/

# Copy built application
COPY dist/ ./dist/
COPY src/ ./src/

# Copy any additional resources
COPY resources/ ./resources/ 2>/dev/null || true

# Create non-root user
RUN addgroup -g 1001 -S nodejs && \
    adduser -S docops -u 1001

# Set permissions
RUN chown -R docops:nodejs /app

# Switch to non-root user
USER docops

# Health check endpoint
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD node dist/index.cjs --health-check || exit 1

# Run the application
CMD ["node", "dist/index.cjs"]