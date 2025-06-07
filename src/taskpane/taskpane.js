/* global document, fetch, console, LM_API_URL */
const DEFAULT_LM_API_URL = "http://localhost:1234/v1/models";
// LM_API_URL is replaced at build time via Webpack's DefinePlugin.
const MODELS_URL =
  typeof LM_API_URL !== "undefined" ? LM_API_URL : DEFAULT_LM_API_URL;

document.addEventListener("DOMContentLoaded", () => {
  const modelDropdown = document.getElementById("modelDropdown");
  const modelError = document.getElementById("modelError");
  const modelSpinner = document.getElementById("spinner");

  const RETRY_DELAY_MS = 5000;
  let retryTimeout = null;

  async function loadModels() {
    // clear any pending retry since we're attempting a fetch now
    if (retryTimeout) {
      clearTimeout(retryTimeout);
      retryTimeout = null;
    }

    modelSpinner.style.display = "inline-block";
    modelError.textContent = "";
    modelDropdown.disabled = true;

    try {
      const res = await fetch(MODELS_URL);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = await res.json();
      if (!Array.isArray(data.models))
        throw new Error("Unexpected model response");

      modelDropdown.innerHTML = "";
      data.models.forEach((model) => {
        const opt = document.createElement("option");
        opt.value = model.id;
        opt.textContent = model.id;
        modelDropdown.appendChild(opt);
      });

      modelDropdown.disabled = false;
      modelSpinner.style.display = "none";
    } catch (err) {
      console.error(err);
      modelDropdown.innerHTML = `<option disabled>Error loading models</option>`;
      modelError.textContent =
        "LM Studio is not running. Retrying connection...";
      modelDropdown.disabled = true;

      // schedule another attempt to fetch models
      retryTimeout = setTimeout(loadModels, RETRY_DELAY_MS);
    }
  }

  loadModels();
});
