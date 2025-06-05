const express = require("express");
const cors = require("cors");
const { createProxyMiddleware } = require("http-proxy-middleware");

const app = express();
const PORT = 4321;

app.use(cors());

app.use(
  "/v1",
  createProxyMiddleware({
    target: "http://127.0.0.1:1234",
    changeOrigin: true,
    pathRewrite: {
      "^/v1": "/v1"
    }
  })
);

app.listen(PORT, () => {
  console.log(`✅ Proxy running at http://localhost:${PORT}/v1`);
});
