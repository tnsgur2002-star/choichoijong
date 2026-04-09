const form = document.getElementById("upload-form");
const input = document.getElementById("file-input");
const label = document.getElementById("file-label");
const dropzone = document.getElementById("dropzone");
const statusEl = document.getElementById("status");
const submitBtn = document.getElementById("submit-btn");

function setStatus(message, level = "") {
  statusEl.textContent = message;
  statusEl.className = `status ${level}`.trim();
}

function updateFileLabel() {
  const f = input.files && input.files[0];
  label.textContent = f ? `${f.name} (${Math.ceil(f.size / 1024)}KB)` : "파일을 드래그하거나 클릭해서 선택하세요";
}

function selectDroppedFile(file) {
  const dt = new DataTransfer();
  dt.items.add(file);
  input.files = dt.files;
  updateFileLabel();
}

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("dragging");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragging");
});

dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("dragging");
  if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
    selectDroppedFile(e.dataTransfer.files[0]);
  }
});

input.addEventListener("change", updateFileLabel);

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  const file = input.files && input.files[0];
  if (!file) {
    setStatus("파일을 먼저 선택해주세요.", "warn");
    return;
  }

  const allowed = [".hwp", ".hwpx", ".pdf"];
  const ext = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
  if (!allowed.includes(ext)) {
    setStatus("지원 형식이 아닙니다. .hwp, .hwpx, .pdf 만 가능합니다.", "warn");
    return;
  }

  submitBtn.disabled = true;
  setStatus("변환 중입니다. 잠시만 기다려주세요...");

  try {
    const fd = new FormData();
    fd.append("document", file);

    const res = await fetch("/api/convert", {
      method: "POST",
      body: fd,
    });

    if (!res.ok) {
      let message = "변환에 실패했습니다.";
      try {
        const j = await res.json();
        if (j && j.message) message = j.message;
      } catch (_) {
        // no-op
      }
      throw new Error(message);
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const fallback = `${file.name.replace(/\.[^.]+$/, "")}_정리.xlsx`;
    const disposition = res.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename\*=UTF-8''([^;]+)|filename="?([^"]+)"?/i);
    const rawName = decodeURIComponent((match && (match[1] || match[2])) || fallback);
    a.href = url;
    a.download = rawName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus("변환 완료. 다운로드를 시작했습니다.", "ok");
  } catch (err) {
    setStatus(err.message || "알 수 없는 오류가 발생했습니다.", "warn");
  } finally {
    submitBtn.disabled = false;
  }
});
