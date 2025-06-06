// ExcelJax: Task Pane Feature Stubs for all required features

let LMStudioAvailable = false;

document.addEventListener("DOMContentLoaded", () => {
  // Model loading and fallback
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
      if (!Array.isArray(data.models)) throw new Error("Unexpected model response");
      modelDropdown.innerHTML = "";
      data.models.forEach(model => {
        const opt = document.createElement("option");
        opt.value = model.id;
        opt.textContent = model.id;
        modelDropdown.appendChild(opt);
      });
      modelDropdown.disabled = false;
      LMStudioAvailable = true;
    } catch (err) {
      modelDropdown.innerHTML = `<option disabled>Error loading models</option>`;
      modelError.textContent = "LM Studio is not running. Features requiring LLMs are disabled.";
      modelDropdown.disabled = true;
      LMStudioAvailable = false;
    } finally {
      modelSpinner.style.display = "none";
      // Optionally disable all LLM-dependent buttons here
      setLLMFeatureState(LMStudioAvailable);
    }
  }

  function setLLMFeatureState(enabled) {
    // Disable/enable all buttons that use LLM
    document.querySelectorAll("button:not(#resetBtn):not(#topChatSendBtn)").forEach(btn => {
      btn.disabled = !enabled;
    });
  }

  // Top Chat Box (general chat, always available)
  document.getElementById("topChatSendBtn").addEventListener("click", async () => {
    const input = document.getElementById("topChatInput");
    const output = document.getElementById("topChatOutput");
    const spinner = document.getElementById("topChatSpinner");
    const prompt = input.value.trim();
    if (!prompt) return;
    output.textContent = ""; spinner.style.display = "inline-block";
    try {
      // You can send to LM Studio here or a cloud endpoint
      const response = await fetch("http://localhost:1234/v1/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: modelDropdown.value, prompt })
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

  // --- Formula/Function Tools section ---
  document.getElementById("replaceFormulaBtn").addEventListener("click", async () => {
    await handleExcelFeature("replaceFormula");
  });
  document.getElementById("explainFormulaBtn").addEventListener("click", async () => {
    await handleExcelFeature("explainFormula");
  });
  document.getElementById("fixFormulaBtn").addEventListener("click", async () => {
    await handleExcelFeature("fixFormula");
  });
  document.getElementById("simplifyFormulaBtn").addEventListener("click", async () => {
    await handleExcelFeature("simplifyFormula");
  });

  // --- Insights & Summarization section ---
  document.getElementById("summarizeDataBtn").addEventListener("click", async () => {
    await handleExcelFeature("summarizeData");
  });
  document.getElementById("keyTakeawaysBtn").addEventListener("click", async () => {
    await handleExcelFeature("keyTakeaways");
  });
  document.getElementById("flagUnusualBtn").addEventListener("click", async () => {
    await handleExcelFeature("flagUnusual");
  });
  document.getElementById("suggestVizBtn").addEventListener("click", async () => {
    await handleExcelFeature("suggestViz");
  });

  // --- Cleanup & Formatting section ---
  document.getElementById("cleanValuesBtn").addEventListener("click", async () => {
    await handleExcelFeature("cleanValues");
  });
  document.getElementById("removeDupesBtn").addEventListener("click", async () => {
    await handleExcelFeature("removeDupes");
  });
  document.getElementById("condFormatBtn").addEventListener("click", async () => {
    await handleExcelFeature("condFormat");
  });
  document.getElementById("reformatHeadersBtn").addEventListener("click", async () => {
    await handleExcelFeature("reformatHeaders");
  });

  // --- Table and Text Tools section ---
  document.getElementById("describeTableBtn").addEventListener("click", async () => {
    await handleExcelFeature("describeTable");
  });
  document.getElementById("autofillDescBtn").addEventListener("click", async () => {
    await handleExcelFeature("autofillDesc");
  });
  document.getElementById("dropdownValuesBtn").addEventListener("click", async () => {
    await handleExcelFeature("dropdownValues");
  });
  document.getElementById("textToTableBtn").addEventListener("click", async () => {
    await handleExcelFeature("textToTable");
  });

  // --- Advanced Excel Use Cases section ---
  document.getElementById("suggestPivotBtn").addEventListener("click", async () => {
    await handleExcelFeature("suggestPivot");
  });
  document.getElementById("nameRangeBtn").addEventListener("click", async () => {
    await handleExcelFeature("nameRange");
  });
  document.getElementById("addHelperColBtn").addEventListener("click", async () => {
    await handleExcelFeature("addHelperCol");
  });
  document.getElementById("predictMissingBtn").addEventListener("click", async () => {
    await handleExcelFeature("predictMissing");
  });

  // --- Language and Communication section ---
  document.getElementById("translateBtn").addEventListener("click", async () => {
    await handleExcelFeature("translate");
  });
  document.getElementById("writeSummaryBtn").addEventListener("click", async () => {
    await handleExcelFeature("writeSummary");
  });
  document.getElementById("addCommentsBtn").addEventListener("click", async () => {
    await handleExcelFeature("addComments");
  });

  // --- Reset button ---
  document.getElementById("resetBtn").addEventListener("click", () => {
    // Clear all inputs, outputs, errors, and spinners
    document.querySelectorAll("input, textarea").forEach(el => el.value = "");
    document.querySelectorAll(".result-box").forEach(el => el.textContent = "");
    document.querySelectorAll(".error-msg").forEach(el => el.textContent = "");
    document.querySelectorAll(".spinner").forEach(el => el.style.display = "none");
  });

  // --- Feature handler stub ---
  async function handleExcelFeature(feature) {
    // Show spinner, clear outputs
    const spinner = document.getElementById(feature + "Spinner");
    const output = document.getElementById(feature + "Output");
    if (spinner) spinner.style.display = "inline-block";
    if (output) output.textContent = "";
    globalError.textContent = "";

    try {
      // TODO: Use Office.js to get selection/input
      // TODO: Send to LM Studio as needed, handle response

      // Example: Show placeholder result
      output.textContent = `[${feature}] - Functionality coming soon`;

      // For real use, you will want to:
      // 1. Get selected cells/range via Office.js
      // 2. Prepare prompt/system message as needed
      // 3. Send to LM Studio (fetch POST)
      // 4. Parse response, update Excel sheet and/or show output in result box

    } catch (err) {
      globalError.textContent = `${feature}: ${err.message}`;
    } finally {
      if (spinner) spinner.style.display = "none";
    }
  }

  // Initial model load
  loadModels();
});
