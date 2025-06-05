/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = analyzeSelectedCells;
  }
});

async function sendToLLM(prompt) {
  const response = await fetch("http://localhost:4321/v1/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "meta-llama-3.1-8b-instruct",  // Make sure this matches your LM Studio model ID
      prompt: prompt,
      max_tokens: 200,
      temperature: 0.7
    })
  });

  const data = await response.json();

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

      const result = await sendToLLM(prompt);

      // Output result safely to one cell just below the selected range
      range.getOffsetRange(values.length + 1, 0).getCell(0, 0).values = [[result]];
      await context.sync();
    });
  } catch (error) {
    console.error("❌ Error in analyzeSelectedCells:", error);
  }
}
