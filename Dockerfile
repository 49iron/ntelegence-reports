FROM node:20-slim

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

# Render sets PORT automatically; default 4000
ENV PORT=4000

# Default: run the HTTP server
# To run scheduler instead, set START_MODE=scheduler in Render env vars
CMD ["sh", "-c", "if [ \"$START_MODE\" = 'scheduler' ]; then node scheduler.js; else node server.js; fi"]
