# Use official Node.js runtime as base image with fixed hash for security
FROM node:18-alpine@sha256:c698ffe060d198dcc6647be78ea1683363f12d5d507dc5ec9855f1c55966ffdd

# Set working directory
WORKDIR /app

# Install system dependencies for headless Chrome
RUN apk add --no-cache \
    chromium \
    nss \
    freetype \
    freetype-dev \
    harfbuzz \
    ca-certificates \
    ttf-liberation \
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

# Install pnpm
RUN npm install -g pnpm

# Install dependencies
RUN pnpm install --frozen-lockfile --prod

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