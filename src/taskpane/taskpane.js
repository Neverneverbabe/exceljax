/* global console, document, Excel, Office, fetch, localStorage */

const defaultConfig = {
  endpoint: "http://localhost:4321/v1/completions",
  model: "meta-llama-3.1-8b-instruct",
};

let config = { ...defaultConfig, ...JSON.parse(localStorage.getItem("llmConfig") || "{}") };
const conversationHistory = [];
let statusElem;

function saveConfig() {
  localStorage.setItem("llmConfig", JSON.stringify(config));
}

function updateInputsFromConfig() {
  document.getElementById("endpointInput").value = config.endpoint;
  document.getElementById("modelInput").value = config.model;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("chatBtn").onclick = chatWithAI;
    document.getElementById("analyzeSelectionBtn").onclick = analyzeSelectedCells;
    document.getElementById("analyzeSheetBtn").onclick = analyzeEntireSheet;
    document.getElementById("suggestFormulaBtn").onclick = suggestFormulaForSelectedCells;
    document.getElementById("improveFormulaBtn").onclick = improveExistingFormula;
    document.getElementById("createVisualsBtn").onclick = createVisualsBasedOnData;
    document.getElementById("saveSettingsBtn").onclick = () => {
      config.endpoint =
        document.getElementById("endpointInput").value.trim() || defaultConfig.endpoint;
      config.model = document.getElementById("modelInput").value.trim() || defaultConfig.model;
      saveConfig();
    };
    statusElem = document.getElementById("status");
    updateInputsFromConfig();
  }
});

async function sendToLLM(prompt) {
  const response = await fetch(config.endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: config.model,
      prompt: prompt,
      max_tokens: 200,
      temperature: 0.7,
    }),
  }).catch((err) => {
    console.error("❌ Fetch error:", err);
    throw err;
  });

  const data = await response.json().catch((err) => {
    console.error("❌ Invalid JSON:", err);
    throw err;
  });

  // Handle both completion and chat-style formats
  if (data.choices?.[0]?.text) {
    return data.choices[0].text.trim();
  } else if (data.choices?.[0]?.message?.content) {
    return data.choices[0].message.content.trim();
  } else {
    return "⚠ No valid response format found from LLM.";
  }
}

async function analyzeSelectedCells() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const values = range.values;
      const prompt = `Analyze the following Excel data and provide summaries, patterns, or suggestions:\n\n${JSON.stringify(values)}`;
      statusElem.textContent = "";
      const result = await sendToLLM(prompt);

      // Output result safely to one cell just below the selected range
      range.getOffsetRange(values.length + 1, 0).getCell(0, 0).values = [[result]];
      await context.sync();
    });
  } catch (error) {
    console.error("❌ Error in analyzeSelectedCells:", error);
    if (statusElem) {
      statusElem.textContent = "Error communicating with LLM";
    }
  }
}

async function chatWithAI() {
  try {
    const input = document.getElementById("promptInput").value.trim();
    if (!input) {
      return;
    }
    const output = document.getElementById("chatOutput");
    output.textContent = "Thinking...";
    statusElem.textContent = "";
    conversationHistory.push(`User: ${input}`);
    const result = await sendToLLM(conversationHistory.join("\n"));
    conversationHistory.push(`Assistant: ${result}`);
    output.textContent = result;
  } catch (error) {
    console.error("❌ Error in chatWithAI:", error);
    if (statusElem) {
      statusElem.textContent = "Error communicating with LLM";
    }
  }
}

async function analyzeEntireSheet() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load("values");
      await context.sync();

      const prompt = `Summarize the following worksheet data:\n\n${JSON.stringify(used.values)}`;
      statusElem.textContent = "";
      const result = await sendToLLM(prompt);

      const summarySheet = context.workbook.worksheets.add("LLM Summary");
      summarySheet.getRange("A1").values = [[result]];
      summarySheet.activate();
      await context.sync();
    });
  } catch (error) {
    console.error("❌ Error in analyzeEntireSheet:", error);
    if (statusElem) {
      statusElem.textContent = "Error communicating with LLM";
    }
  }
}

async function suggestFormulaForSelectedCells() {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getSelectedRange();
      cell.load("address");
      await context.sync();

      const description = document.getElementById("promptInput").value.trim();
      if (!description) {
        return;
      }

      const prompt = `Provide an Excel formula only for: ${description}`;
      statusElem.textContent = "";
      const formula = await sendToLLM(prompt);
      cell.formulas = [[formula]];
      await context.sync();
    });
  } catch (error) {
    console.error("❌ Error in suggestFormulaForSelectedCells:", error);
    if (statusElem) {
      statusElem.textContent = "Error communicating with LLM";
    }
  }
}

async function improveExistingFormula() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("formulas");
      await context.sync();

      const formulas = range.formulas;
      const prompt = `Improve or simplify these formulas. Return them in the same order separated by new lines:\n${JSON.stringify(formulas)}`;
      statusElem.textContent = "";
      const result = await sendToLLM(prompt);
      const lines = result.split(/\r?\n/).filter((l) => l.trim().length > 0);

      const rows = formulas.length;
      const cols = formulas[0].length;
      if (lines.length === rows * cols) {
        const newFormulas = [];
        let idx = 0;
        for (let r = 0; r < rows; r++) {
          newFormulas[r] = [];
          for (let c = 0; c < cols; c++) {
            newFormulas[r][c] = lines[idx++] || formulas[r][c];
          }
        }
        range.formulas = newFormulas;
      } else if (lines.length === 1 && rows === 1 && cols === 1) {
        range.formulas = [[lines[0]]];
      }
      await context.sync();
    });
  } catch (error) {
    console.error("❌ Error in improveExistingFormula:", error);
    if (statusElem) {
      statusElem.textContent = "Error communicating with LLM";
    }
  }
}

async function createVisualsBasedOnData() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const prompt = `Suggest the best chart type for this data: ${JSON.stringify(range.values)}`;
      statusElem.textContent = "";
      const suggestion = await sendToLLM(prompt);
      let chartType = "ColumnClustered";
      const text = suggestion.toLowerCase();
      if (text.includes("line")) chartType = "Line";
      else if (text.includes("pie")) chartType = "Pie";
      else if (text.includes("bar")) chartType = "BarClustered";
      else if (text.includes("scatter")) chartType = "XYScatter";

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.columns);
      chart.setPosition(
        range.getOffsetRange(range.rowCount + 1, 0),
        range.getOffsetRange(range.rowCount + 15, range.columnCount)
      );
      await context.sync();
    });
  } catch (error) {
    console.error("❌ Error in createVisualsBasedOnData:", error);
    if (statusElem) {
      statusElem.textContent = "Error communicating with LLM";
    }
  }
}
