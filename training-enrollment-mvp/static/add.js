const sessionResult = document.getElementById("session-result");
const importResult = document.getElementById("import-result");
const importSessionHint = document.getElementById("import-session-hint");
const pageTitle = document.getElementById("page-title");
const saveSessionBtn = document.getElementById("save-session");

let currentSessionId = null;

function showResult(element, message, isError = false) {
  element.classList.remove("hidden");
  element.classList.toggle("error", isError);
  element.innerHTML = message;
}

function getQuerySessionId() {
  const url = new URL(window.location.href);
  const sessionId = url.searchParams.get("session_id");
  if (!sessionId || !/^\d+$/.test(sessionId)) {
    return null;
  }
  return Number(sessionId);
}

function fillSessionForm(data) {
  document.getElementById("title").value = data.title || "";
  document.getElementById("start_date").value = data.start_date || "";
  document.getElementById("end_date").value = data.end_date || "";
  document.getElementById("location_text").value = data.location_text || "";
}

async function handleResponse(response) {
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "请求失败");
  }
  return payload.data;
}

async function loadSessionForEdit() {
  const querySessionId = getQuerySessionId();
  if (!querySessionId) {
    return;
  }

  try {
    const response = await fetch(`/api/session/${querySessionId}`);
    const data = await handleResponse(response);
    currentSessionId = data.session_id;
    fillSessionForm(data);
    pageTitle.textContent = `编辑培训 #${currentSessionId}`;
    saveSessionBtn.textContent = "保存修改";
    importSessionHint.textContent = `当前绑定期次 session_id: ${currentSessionId}`;
  } catch (error) {
    showResult(sessionResult, `加载期次失败：${error.message}`, true);
  }
}

saveSessionBtn.addEventListener("click", async () => {
  const formData = new FormData();
  formData.append("title", document.getElementById("title").value.trim());
  formData.append("start_date", document.getElementById("start_date").value);
  formData.append("end_date", document.getElementById("end_date").value);
  formData.append("location_text", document.getElementById("location_text").value.trim());

  const noticeFile = document.getElementById("notice_file").files[0];
  if (noticeFile) {
    formData.append("notice_file", noticeFile);
  }

  try {
    let response;
    if (currentSessionId) {
      formData.append("session_id", String(currentSessionId));
      response = await fetch("/api/session/update", { method: "POST", body: formData });
    } else {
      response = await fetch("/api/session/create", { method: "POST", body: formData });
    }

    const data = await handleResponse(response);
    currentSessionId = data.session_id;
    saveSessionBtn.textContent = "保存修改";
    pageTitle.textContent = `编辑培训 #${currentSessionId}`;
    showResult(sessionResult, `期次保存成功，session_id: ${currentSessionId}`);
    importSessionHint.textContent = `当前绑定期次 session_id: ${currentSessionId}`;
  } catch (error) {
    showResult(sessionResult, error.message, true);
  }
});

document.getElementById("import-enrollment").addEventListener("click", async () => {
  if (!currentSessionId) {
    showResult(importResult, "请先完成第一步创建期次，再导入报名 Excel。", true);
    return;
  }

  const excelFile = document.getElementById("excel_file").files[0];
  if (!excelFile) {
    showResult(importResult, "请先选择报名 Excel 文件。", true);
    return;
  }

  const allowedExcelSuffix = [".xlsx", ".xls", ".xlsm", ".xltx", ".xltm"];
  const lowerName = excelFile.name.toLowerCase();
  if (!allowedExcelSuffix.some((suffix) => lowerName.endsWith(suffix))) {
    showResult(importResult, "仅支持 Excel 文件（.xlsx/.xls/.xlsm/.xltx/.xltm）。", true);
    return;
  }

  const formData = new FormData();
  formData.append("excel_file", excelFile);
  formData.append("session_id", String(currentSessionId));

  try {
    const response = await fetch("/api/enrollment/import", { method: "POST", body: formData });
    const data = await handleResponse(response);
    const exceptionItems = (data.exceptions || []).map((item) =>
      `<li>Sheet: ${item.sheet || ""} 行: ${item.row || ""} 原因: ${item.reason}</li>`
    );
    const html = `
      <p>Sheet 数: ${data.sheet_count}</p>
      <p>有效行数: ${data.valid_rows}</p>
      <p>新增 person: ${data.new_person_count}</p>
      <p>新增 enrollment: ${data.new_enrollment_count}</p>
      <p>异常行:</p>
      <ul>${exceptionItems.join("") || "<li>无</li>"}</ul>
    `;
    showResult(importResult, html);
  } catch (error) {
    showResult(importResult, error.message, true);
  }
});

loadSessionForEdit();
