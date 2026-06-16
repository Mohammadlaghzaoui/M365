# Container for the WorkPilot Migration Agent — deploy to any cloud (Render,
# Railway, Fly.io, Koyeb, Azure Container Apps, Google Cloud Run, ...).
# Runs the read-only device-code discovery server-side. Nothing local.
FROM node:20-alpine

WORKDIR /app
COPY package*.json ./
RUN npm install --omit=dev
COPY . .

# The platform injects PORT; bind to all interfaces so the proxy can reach us.
ENV HOST=0.0.0.0
EXPOSE 8787

# AGENT_API_KEY and ALLOWED_ORIGINS must be supplied as environment variables.
CMD ["npm", "start"]
