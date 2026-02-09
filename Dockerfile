# ruleid: dockerfile-node-18-detected
FROM node:18-alpine

WORKDIR /app

# Install curl for healthcheck
RUN apk add --no-cache curl

COPY sns-subscriber/package*.json ./
RUN npm install

COPY sns-subscriber/ .

EXPOSE 9000

CMD ["node", "subscriber.js"]

# ok: dockerfile-node-18-detected
# node:18-alpine
