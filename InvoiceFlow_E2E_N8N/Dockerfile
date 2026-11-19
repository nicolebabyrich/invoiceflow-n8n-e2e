FROM n8n-base:2.4.6
USER root
RUN npm install -g xlsx && \
    ln -sfn /opt/nodejs/node-v22.21.1/lib/node_modules/exceljs /usr/local/lib/node_modules/n8n/node_modules/exceljs && \
    ln -sfn /opt/nodejs/node-v22.21.1/lib/node_modules/xlsx /usr/local/lib/node_modules/n8n/node_modules/xlsx
ENV NODE_PATH=/opt/nodejs/node-v22.21.1/lib/node_modules
USER node
