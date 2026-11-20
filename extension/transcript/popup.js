const captureButton = document.querySelector("#capture");
const copyButton = document.querySelector("#copy");
const formatSelect = document.querySelector("#format");
const status = document.querySelector("#status");
const diagnostics = document.querySelector("#diagnostics");
const diagnosticsOutput = document.querySelector("#diagnostics-output");

function setStatus(message, state = "working") {
  status.textContent = message;
  status.dataset.state = state;
}

captureButton.addEventListener("click", () => capture("download"));
copyButton.addEventListener("click", () => capture("copy"));

async function capture(destination) {
  captureButton.disabled = true;
  copyButton.disabled = true;
  formatSelect.disabled = true;
  diagnostics.hidden = true;
  diagnosticsOutput.textContent = "";
  setStatus("Capturing the transcript. Keep the Teams tab open.");

  try {
    const response = await chrome.runtime.sendMessage({
      type: "CAPTURE_TEAMS_TRANSCRIPT",
      format: formatSelect.value,
      destination,
    });

    if (!response?.ok) {
      if (response?.diagnostics) {
        diagnosticsOutput.textContent = JSON.stringify(response.diagnostics, null, 2);
        diagnostics.hidden = false;
      }
      throw new Error(response?.error || "The capture did not complete.");
    }

    if (destination === "copy") {
      try {
        await navigator.clipboard.writeText(response.content);
      } catch {
        throw new Error("Could not copy the transcript to the clipboard. Keep this popup open or use Capture and download.");
      }
      setStatus(`Copied ${response.count} transcript entries to the clipboard.`, "success");
    } else {
      setStatus(`Downloaded ${response.count} transcript entries as ${response.filename}.`, "success");
    }
  } catch (error) {
    setStatus(error.message || String(error), "error");
  } finally {
    captureButton.disabled = false;
    copyButton.disabled = false;
    formatSelect.disabled = false;
  }
}
