function showMessage(text) {
  document.getElementById("message").textContent = text;
  document.getElementById("error").textContent = "";
  setTimeout(() => {
    document.getElementById("message").textContent = "";
  }, 2000);
}

function showError(text) {
  document.getElementById("error").textContent = text;
  document.getElementById("message").textContent = "";
}

function copyToClipboard(code) {
  navigator.clipboard.writeText(code).then(() => {
    showMessage("コピーしました！");
  }).catch(() => {
    showError("コピーに失敗しました");
  });
}

function generateCode(selectedText, style) {
  return `<span style="${style}">${selectedText}</span>`;
}

function getSelectedTextAndGenerate(styleCallback) {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const tabId = tabs[0].id;
    chrome.tabs.sendMessage(tabId, { action: "getSelectedText" }, (response) => {
      if (chrome.runtime.lastError || !response) {
        showError("ページを再読み込みしてお試しください");
        return;
      }
      const text = response.text;
      if (!text) {
        showError("文章が選択されていません");
        return;
      }
      const code = generateCode(text, styleCallback());
      copyToClipboard(code);
    });
  });
}

document.getElementById("btn-larger").addEventListener("click", () => {
  getSelectedTextAndGenerate(() => "font-size: 1.2em;");
});

document.getElementById("btn-smaller").addEventListener("click", () => {
  getSelectedTextAndGenerate(() => "font-size: 0.8em;");
});

document.getElementById("btn-color").addEventListener("click", () => {
  const color = document.getElementById("pick-color").value;
  getSelectedTextAndGenerate(() => `color: ${color};`);
});

document.getElementById("btn-marker").addEventListener("click", () => {
  const color = document.getElementById("pick-marker").value;
  getSelectedTextAndGenerate(() => `background-color: ${color};`);
});
