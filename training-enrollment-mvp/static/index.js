const statsResult = document.getElementById("stats-result");
const historyList = document.getElementById("history-list");

function renderTop5(items) {
  if (!items || items.length === 0) {
    return "<p>无</p>";
  }
  const rows = items.map((item) => `<tr><td>${item.phone_norm}</td><td>${item.name || ""}</td><td>${item.count}</td></tr>`);
  return `<table><thead><tr><th>手机号</th><th>姓名</th><th>次数</th></tr></thead><tbody>${rows.join("")}</tbody></table>`;
}

async function handleResponse(response) {
  const payload = await response.json();
  if (!payload.ok) {
    throw new Error(payload.error || "请求失败");
  }
  return payload.data;
}

async function fetchStats() {
  const year = document.getElementById("year").value.trim();
  if (!year) {
    statsResult.innerHTML = '<p class="error">请输入年份。</p>';
    return;
  }

  try {
    const response = await fetch(`/api/stats/year?year=${encodeURIComponent(year)}`);
    const data = await handleResponse(response);
    statsResult.innerHTML = `
      <p>参训人次: ${data.total_enrollments}</p>
      <p>参训人数: ${data.total_people}</p>
      <p>复训人数: ${data.repeat_people}</p>
      <p>Top5:</p>
      ${renderTop5(data.top5)}
    `;
  } catch (error) {
    statsResult.innerHTML = `<p class="error">${error.message}</p>`;
  }
}

async function fetchHistory() {
  try {
    const response = await fetch("/api/session/history");
    const sessions = await handleResponse(response);
    if (!sessions || sessions.length === 0) {
      historyList.innerHTML = "<p>暂无历史培训记录。</p>";
      return;
    }

    historyList.innerHTML = sessions
      .map(
        (item) => `
          <div class="history-item">
            <div><strong>#${item.session_id} ${item.title || "未命名培训"}</strong></div>
            <div>时间：${item.start_date || ""} ~ ${item.end_date || ""}</div>
            <div>地点：${item.location_text || ""}</div>
            <div>报名人次：${item.enrollment_count}</div>
            <div>创建时间：${item.created_at || ""}</div>
          </div>
        `
      )
      .join("");
  } catch (error) {
    historyList.innerHTML = `<p class="error">加载历史培训失败：${error.message}</p>`;
  }
}

document.getElementById("fetch-stats").addEventListener("click", fetchStats);
document.getElementById("export-year").addEventListener("click", () => {
  const year = document.getElementById("year").value.trim();
  if (!year) {
    statsResult.innerHTML = '<p class="error">请输入年份再导出。</p>';
    return;
  }
  window.location.href = `/api/export/year?year=${encodeURIComponent(year)}`;
});
document.getElementById("refresh-history").addEventListener("click", fetchHistory);

fetchHistory();
