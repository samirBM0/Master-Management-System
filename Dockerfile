FROM node:20-alpine

WORKDIR /app

# Native build toolchain required to compile better-sqlite3 on musl/Alpine
# (bcryptjs is pure-JS and needs no compilation).
RUN apk add --no-cache python3 make g++

COPY package*.json bun.lock* ./
RUN npm install

COPY . .

RUN npm run build

# Persist the SQLite database (seeded users + revoked tokens) in /app/data,
# which is bind-mounted to ./data by docker-compose.yml.
VOLUME ["/app/data"]
ENV DATA_DIR=/app/data

EXPOSE 3005

CMD ["npx", "tsx", "server.ts"]
