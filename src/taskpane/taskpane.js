document.addEventListener("DOMContentLoaded", () => {
  const modelDropdown = document.getElementById("modelDropdown");
  const modelError = document.getElementById("modelError");
  const modelSpinner = document.getElementById("spinner");

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
    } catch (err) {
      modelDropdown.innerHTML = `<option disabled>Error loading models</option>`;
      modelError.textContent =
        "LM Studio is not running. Features requiring LLMs are disabled.";
      modelDropdown.disabled = true;
    } finally {
      modelSpinner.style.display = "none";
    }
  }

  loadModels();
});
