# ExcelJax Excel Add-in

This repository contains a sample Excel add‑in that communicates with a local language model. The add‑in sends all OpenAI compatible requests through a small proxy server and expects a backend LLM service to be running locally.

## Starting the local services

1. **Start your LLM server**
   
   Run your preferred backend LLM service so that it listens on `http://127.0.0.1:1234` and exposes an OpenAI compatible `/v1` endpoint. For example, LM Studio can be started with:
   
   ```bash
   lm-studio --server --port 1234 --model meta-llama-3.1-8b-instruct
   ```

2. **Start the proxy**
   
   From the project directory, start the proxy server:
   
   ```bash
   node proxy-server.js
   ```

   The proxy listens on `http://localhost:4321/v1` and forwards all requests to the LLM service.

### Run them automatically

To use the add‑in offline without typing commands each time, create system services for the LLM server and `proxy-server.js` (e.g., systemd units on Linux or background services on Windows). Once configured, both processes will launch automatically at startup, ensuring the add‑in can reach the language model without manual steps.
