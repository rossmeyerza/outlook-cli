const captureButton = document.querySelector("#capture");
const formatSelect = document.querySelector("#format");
const status = document.querySelector("#status");
const diagnostics = document.querySelector("#diagnostics");
const diagnosticsOutput = document.querySelector("#diagnostics-output");

function setStatus(message, state = "working") {
  status.textContent = message;
  status.dataset.state = state;
}

captureButton.addEventListener("click", async () => {
  captureButton.disabled = true;
  formatSelect.disabled = true;
  diagnostics.hidden = true;
  diagnosticsOutput.textContent = "";
  setStatus("Capturing the transcript. Keep the Teams tab open.");

  try {
    const response = await chrome.runtime.sendMessage({
      type: "CAPTURE_TEAMS_TRANSCRIPT",
      format: formatSelect.value,
    });

    if (!response?.ok) {
      if (response?.diagnostics) {
        diagnosticsOutput.textContent = JSON.stringify(response.diagnostics, null, 2);
        diagnostics.hidden = false;
      }
      throw new Error(response?.error || "The capture did not complete.");
    }

    setStatus(
      `Downloaded ${response.count} transcript entries as ${response.filename}.`,
      "success",
    );
  } catch (error) {
    setStatus(error.message || String(error), "error");
  } finally {
    captureButton.disabled = false;
    formatSelect.disabled = false;
  }
});
