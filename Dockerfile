FROM node:20-alpine

WORKDIR /app

COPY package.json package-lock.json bun.lock* ./
RUN npm install

COPY . .

RUN npm run build

EXPOSE 3005

CMD ["npx", "tsx", "server.ts"]
