# ===== WorkPilot all-in-one server =====
# Builds the portal and serves it together with the agent API + executor.
# Note: on-prem AD/Exchange PowerShell needs a Windows host; in Linux/Docker the
# agent runs the cloud (Microsoft Graph) paths. For full hybrid, run the
# Windows service install (see scripts/install-windows.ps1).

# ---- build the portal ----
FROM node:20-alpine AS build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
RUN npm run build:server

# ---- runtime: agent + bundled portal ----
FROM node:20-alpine
WORKDIR /app/agent
COPY agent/package*.json ./
RUN npm ci --omit=dev
COPY agent/ ./
COPY --from=build /app/agent/public ./public
ENV PORT=8787
EXPOSE 8787
CMD ["npm", "start"]
