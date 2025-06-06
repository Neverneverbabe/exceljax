/* global fetch, document */
async function callLocalLLM(modelId, promptText) {
  const response = await fetch("http://localhost:1234/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: modelId,
      messages: [
        { role: "system", content: "You are a helpful assistant running inside Excel." },
        { role: "user", content: promptText },
      ],
      temperature: 0.7,
      max_tokens: -1,
      stream: false,
    }),
  });

  if (!response.ok) {
    throw new Error("Model server error");
  }

  const data = await response.json();
  return data.choices[0].message.content;
}

document.getElementById("submitPrompt").addEventListener("click", async () => {
  const modelId = document.getElementById("modelSelect").value;
  const promptText = document.getElementById("userPrompt").value;
  const outputBox = document.getElementById("responseOutput");

  outputBox.innerText = "Thinking...";

  try {
    const result = await callLocalLLM(modelId, promptText);
    outputBox.innerText = result;
  } catch (err) {
    outputBox.innerText = "❌ Error: " + err.message;
  }
});
