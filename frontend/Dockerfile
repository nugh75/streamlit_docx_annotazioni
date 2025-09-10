# syntax=docker/dockerfile:1

# Build stage
FROM node:20-alpine AS build
WORKDIR /app
# Copy package manifests first for better caching
COPY frontend/package*.json ./
# Install deps (use lockfile if present)
RUN if [ -f package-lock.json ]; then npm ci; else npm install; fi
# Copy the rest of the frontend
COPY frontend .
RUN npm run build

# Runtime stage
FROM nginx:1.27-alpine AS runtime
WORKDIR /usr/share/nginx/html
# Copy built assets
COPY --from=build /app/dist ./
# Nginx config with API proxy
COPY frontend/nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
