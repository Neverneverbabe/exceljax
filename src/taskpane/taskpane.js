/* global console, document, Excel, Office, fetch, localStorage */

const defaultConfig = {
  endpoint: "http://localhost:1234/v1/chat/completions",
  model: "meta-llama-3.1-8b-instruct",
  token: "", // optional auth token
  maxHistory: 10,
};

let config = { ...defaultConfig, ...JSON.parse(localStorage.getItem("llmConfig") || "{}") };
const conversationHistory = [];
let statusElem;
let dialog = null;

function saveConfig() {
  localStorage.setItem("llmConfig", JSON.stringify(config));
}

function updateInputsFromConfig() {
  document.getElementById("endpointInput").value = config.endpoint;
  document.getElementById("modelInput").value = config.model;
  const tokenInput = document.getElementById("tokenInput");
  if (tokenInput) tokenInput.value = config.token;
}

async function fetchWithRetry(url, options, retries = 2) {
  for (let i = 0; i <= retries; i++) {
    try {
      const res = await fetch(url, options);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res;
    } catch (err) {
      if (i === retries) throw err;
    }
  }
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
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("filter-table").onclick = filterTable;
    document.getElementById("sort-table").onclick = sortTable;
    document.getElementById("create-chart").onclick = createChart;
    document.getElementById("freeze-header").onclick = freezeHeader;
    const pivotBtn = document.getElementById("pivot-suggestion");
    if (pivotBtn) pivotBtn.onclick = suggestPivotTable;
    document.getElementById("open-dialog").onclick = openDialog;
    document.getElementById("saveSettingsBtn").onclick = () => {
      config.endpoint =
        document.getElementById("endpointInput").value.trim() || defaultConfig.endpoint;
      config.model = document.getElementById("modelInput").value.trim() || defaultConfig.model;
      const tokenInput = document.getElementById("tokenInput");
      config.token = tokenInput ? tokenInput.value.trim() : "";
      saveConfig();
    };
    statusElem = document.getElementById("status");
    updateInputsFromConfig();
  }
});

async function sendToLLM(prompt) {
  const maxPrompt = 6000;
  if (prompt.length > maxPrompt) {
    prompt = prompt.slice(0, maxPrompt) + "\n[Truncated]";
  }

  const headers = {
    "Content-Type": "application/json",
  };
  if (config.token) {
    headers["Authorization"] = `Bearer ${config.token}`;
  }

  const options = {
    method: "POST",
    headers,
    body: JSON.stringify({
      model: config.model,

      temperature: 0.7,
      stream: false,
    }),
  };

  let response;
  try {
    response = await fetchWithRetry(config.endpoint, options);
  } catch (err) {
    console.error("❌ Fetch error:", err);
    throw new Error(`Failed to reach LLM: ${err.message}`);
  }

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
      statusElem.textContent = error.message || "Error communicating with LLM";
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
    if (conversationHistory.length > config.maxHistory * 2) {
      conversationHistory.splice(0, conversationHistory.length - config.maxHistory * 2);
    }
    const result = await sendToLLM(conversationHistory.join("\n"));
    conversationHistory.push(`Assistant: ${result}`);
    if (conversationHistory.length > config.maxHistory * 2) {
      conversationHistory.splice(0, conversationHistory.length - config.maxHistory * 2);
    }
    output.textContent = result;
  } catch (error) {
    console.error("❌ Error in chatWithAI:", error);
    if (statusElem) {
      statusElem.textContent = error.message || "Error communicating with LLM";
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
      statusElem.textContent = error.message || "Error communicating with LLM";
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
      statusElem.textContent = error.message || "Error communicating with LLM";
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
      statusElem.textContent = error.message || "Error communicating with LLM";
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
      statusElem.textContent = error.message || "Error communicating with LLM";
    }
  }
}

async function createTable() {
  try {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.add("A1:D1", true);
      expensesTable.name = "ExpensesTable";

      expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
      ]);

      expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function filterTable() {
  try {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      const categoryColumn = expensesTable.columns.getItem("Category");
      categoryColumn.load("filter");
      await context.sync();
      categoryColumn.filter.applyValuesFilter(["Education", "Groceries"]);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function sortTable() {
  try {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      const sortFields = [
        {
          key: 1,
          ascending: false,
        },
      ];

      expensesTable.sort.apply(sortFields);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function createChart() {
  try {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
      const dataRange = expensesTable.getDataBodyRange();
      const chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "Auto");
      chart.setPosition("A15", "F30");
      chart.title.text = "Expenses";
      chart.legend.position = "Right";
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
      chart.series.getItemAt(0).name = "Value in \u20AC";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

async function freezeHeader() {
  try {
    await Excel.run(async (context) => {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.freezePanes.freezeRows(1);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

function openDialog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup.html",
    { height: 45, width: 55 },
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}

async function suggestPivotTable() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const prompt = `Suggest a pivot table layout for the following data:\n\n${JSON.stringify(range.values)}`;
      statusElem.textContent = "";
      const result = await sendToLLM(prompt);
      range.getOffsetRange(range.rowCount + 1, 0).getCell(0, 0).values = [[result]];
      await context.sync();
    });
  } catch (error) {
    console.error("❌ Error in suggestPivotTable:", error);
    if (statusElem) {
      statusElem.textContent = error.message || "Error communicating with LLM";
    }
  }
}
