FROM node:24-alpine
WORKDIR /app
COPY --chown=node:node server.js ./
COPY --chown=node:node public ./public
COPY --chown=node:node tests ./tests
RUN mkdir /app/data && chown node:node /app/data
USER node
EXPOSE 8080
CMD ["node", "server.js"]
