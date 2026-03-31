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

// チェックボックスとカラーピッカーの連動
document.getElementById("chk-color").addEventListener("change", (e) => {
  document.getElementById("pick-color").disabled = !e.target.checked;
});

document.getElementById("chk-marker").addEventListener("change", (e) => {
  document.getElementById("pick-marker").disabled = !e.target.checked;
});

// コピーボタン
document.getElementById("btn-copy").addEventListener("click", () => {
  const sizeValue = document.querySelector('input[name="size"]:checked').value;
  const useColor = document.getElementById("chk-color").checked;
  const useMarker = document.getElementById("chk-marker").checked;

  // 何も選ばれていない場合
  if (sizeValue === "none" && !useColor && !useMarker) {
    showError("装飾を一つ以上選んでください");
    return;
  }

  // styleを組み立てる
  const styles = [];
  if (sizeValue === "large") styles.push("font-size: 1.2em;");
  if (sizeValue === "small") styles.push("font-size: 0.8em;");
  if (useColor) {
    const color = document.getElementById("pick-color").value;
    styles.push(`color: ${color};`);
  }
  if (useMarker) {
    const color = document.getElementById("pick-marker").value;
    styles.push(`background-color: ${color};`);
  }

  const styleStr = styles.join(" ");

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
      const code = `<span style="${styleStr}">${text}</span>`;
      copyToClipboard(code);
    });
  });
});
