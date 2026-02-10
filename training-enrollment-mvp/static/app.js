const sessionResult = document.getElementById("session-result");
const importSection = document.getElementById("section-import");
const importResult = document.getElementById("import-result");
const statsResult = document.getElementById("stats-result");

let currentSessionId = null;

function showResult(element, message, isError = false) {
  element.classList.remove("hidden");
  element.classList.toggle("error", isError);
  element.innerHTML = message;
}

function toHtmlList(items) {
  if (!items || items.length === 0) {
    return "<p>无</p>";
  }
  const rows = items.map((item) => {
    return `<tr><td>${item.phone_norm}</td><td>${item.name || ""}</td><td>${item.count}</td></tr>`;
  });
  return `
    <table>
      <thead><tr><th>手机号</th><th>姓名</th><th>次数</th></tr></thead>
      <tbody>${rows.join("")}</tbody>
    </table>
  `;
}

async function handleResponse(response) {
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "请求失败");
  }
  return payload.data;
}

document.getElementById("create-session").addEventListener("click", async () => {
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
    const response = await fetch("/api/session/create", {
      method: "POST",
      body: formData,
    });
    const data = await handleResponse(response);
    currentSessionId = data.session_id;
    showResult(sessionResult, `期次创建成功，session_id: ${currentSessionId}`);
    importSection.classList.remove("hidden");
  } catch (error) {
    showResult(sessionResult, error.message, true);
  }
});

document.getElementById("import-enrollment").addEventListener("click", async () => {
  const excelFile = document.getElementById("excel_file").files[0];
  if (!excelFile) {
    showResult(importResult, "请先选择报名 Excel 文件。", true);
    return;
  }

  const allowedExcelSuffix = [".xlsx", ".xls", ".xlsm", ".xltx", ".xltm"];
  const lowerName = excelFile.name.toLowerCase();
  const validExcel = allowedExcelSuffix.some((suffix) => lowerName.endsWith(suffix));
  if (!validExcel) {
    showResult(importResult, "仅支持 Excel 文件（.xlsx/.xls/.xlsm/.xltx/.xltm）。", true);
    return;
  }

  const formData = new FormData();
  formData.append("excel_file", excelFile);
  if (currentSessionId) {
    formData.append("session_id", currentSessionId);
  }

  try {
    const response = await fetch("/api/enrollment/import", {
      method: "POST",
      body: formData,
    });
    const data = await handleResponse(response);
    const exceptions = data.exceptions || [];
    const exceptionList = exceptions
      .map(
        (item) =>
          `<li>Sheet: ${item.sheet || ""} 行: ${item.row || ""} 原因: ${item.reason}</li>`
      )
      .join("");
    const html = `
      <p>Sheet 数: ${data.sheet_count}</p>
      <p>有效行数: ${data.valid_rows}</p>
      <p>新增 person: ${data.new_person_count}</p>
      <p>新增 enrollment: ${data.new_enrollment_count}</p>
      <p>异常行:</p>
      <ul>${exceptionList || "<li>无</li>"}</ul>
    `;
    showResult(importResult, html);
  } catch (error) {
    showResult(importResult, error.message, true);
  }
});

document.getElementById("fetch-stats").addEventListener("click", async () => {
  const year = document.getElementById("year").value.trim();
  try {
    const response = await fetch(`/api/stats/year?year=${encodeURIComponent(year)}`);
    const data = await handleResponse(response);
    const html = `
      <p>参训人次: ${data.total_enrollments}</p>
      <p>参训人数: ${data.total_people}</p>
      <p>复训人数: ${data.repeat_people}</p>
      <p>Top5:</p>
      ${toHtmlList(data.top5)}
    `;
    showResult(statsResult, html);
  } catch (error) {
    showResult(statsResult, error.message, true);
  }
});

document.getElementById("export-year").addEventListener("click", async () => {
  const year = document.getElementById("year").value.trim();
  if (!year) {
    showResult(statsResult, "请输入年份再导出。", true);
    return;
  }
  window.location.href = `/api/export/year?year=${encodeURIComponent(year)}`;
});
