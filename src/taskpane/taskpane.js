// Polyfill for document/fetch in non-browser environments (for linting/tests only)
if (typeof document === "undefined") {
  global.document = {
    getElementById: () => null,
    querySelectorAll: () => [],
    addEventListener: () => {},
  };
}
if (typeof fetch === "undefined") {
  global.fetch = () => Promise.reject("fetch not available");
}

let LMStudioAvailable = false;

document.addEventListener("DOMContentLoaded", () => {
  // Elements
  const modelDropdown = document.getElementById("modelDropdown");
  const modelError = document.getElementById("modelError");
  const modelSpinner = document.getElementById("spinner");
  const globalError = document.getElementById("globalError");

  async function loadModels() {
    modelSpinner.style.display = "inline-block";
    modelError.textContent = "";
    modelDropdown.disabled = true;
    try {
      const res = await fetch("http://localhost:1234/v1/models");
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
      LMStudioAvailable = true;
    } catch (err) {
      modelDropdown.innerHTML = `<option disabled>Error loading models</option>`;
      modelError.textContent =
        "LM Studio is not running. Features requiring LLMs are disabled.";
      modelDropdown.disabled = true;
      LMStudioAvailable = false;
    } finally {
      modelSpinner.style.display = "none";
      setLLMFeatureState(LMStudioAvailable);
    }
  }

  function setLLMFeatureState(enabled) {
    document
      .querySelectorAll("button:not(#resetBtn):not(#topChatSendBtn)")
      .forEach((btn) => {
        btn.disabled = !enabled;
      });
  }

  // --- Top Chat Box ---
  document
    .getElementById("topChatSendBtn")
    ?.addEventListener("click", async () => {
      const input = document.getElementById("topChatInput");
      const output = document.getElementById("topChatOutput");
      const spinner = document.getElementById("topChatSpinner");
      const prompt = input?.value.trim();
      if (!prompt) return;
      output.textContent = "";
      spinner.style.display = "inline-block";
      try {
        const response = await fetch("http://localhost:1234/v1/completions", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ model: modelDropdown.value, prompt }),
        });
        if (!response.ok) throw new Error("LM Studio not available");
        const data = await response.json();
        output.textContent = data.choices?.[0]?.text ?? "No response";
      } catch (err) {
        output.textContent = "AI not available or LM Studio not running.";
      } finally {
        spinner.style.display = "none";
      }
    });

  // --- Feature Button Handlers (Excel, Insights, Formatting, Table, Advanced, Language) ---
  const featureButtonMap = [
    // Excel
    "replaceFormula",
    "explainFormula",
    "fixFormula",
    "simplifyFormula",
    // Insights
    "summarizeData",
    "keyTakeaways",
    "flagUnusual",
    "suggestViz",
    // Formatting
    "cleanValues",
    "removeDupes",
    "condFormat",
    "reformatHeaders",
    // Table
    "describeTable",
    "autofillDesc",
    "dropdownValues",
    "textToTable",
    // Advanced
    "suggestPivot",
    "nameRange",
    "addHelperCol",
    "predictMissing",
    // Language
    "translate",
    "writeSummary",
    "addComments",
  ];
  featureButtonMap.forEach((feature) => {
    document
      .getElementById(`${feature}Btn`)
      ?.addEventListener("click", async () => {
        await handleExcelFeature(feature);
      });
  });

  // --- Reset button ---
  document.getElementById("resetBtn")?.addEventListener("click", () => {
    document
      .querySelectorAll("input, textarea")
      .forEach((el) => (el.value = ""));
    document
      .querySelectorAll(".result-box")
      .forEach((el) => (el.textContent = ""));
    document
      .querySelectorAll(".error-msg")
      .forEach((el) => (el.textContent = ""));
    document
      .querySelectorAll(".spinner")
      .forEach((el) => (el.style.display = "none"));
  });

  // --- Feature handler stub ---
  async function handleExcelFeature(feature) {
    const spinner = document.getElementById(feature + "Spinner");
    const output = document.getElementById(feature + "Output");
    if (spinner) spinner.style.display = "inline-block";
    if (output) output.textContent = "";
    globalError.textContent = "";

    try {
      // TODO: Use Office.js for Excel integration
      // TODO: Send to LM Studio as needed, handle response

      // Example placeholder
      output.textContent = `[${feature}] - Functionality coming soon`;
    } catch (err) {
      globalError.textContent = `${feature}: ${err.message ?? err}`;
    } finally {
      if (spinner) spinner.style.display = "none";
    }
  }

  // Initial model load
  loadModels();
});
