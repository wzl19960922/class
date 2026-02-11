const statsResult = document.getElementById("stats-result");
const historyList = document.getElementById("history-list");
const todayTaskResult = document.getElementById("today-task-result");
const todayTaskList = document.getElementById("today-task-list");
const logPath = document.getElementById("log-path");
const logContent = document.getElementById("log-content");

const addTrainingModal = document.getElementById("add-training-modal");
const newSessionResult = document.getElementById("new-session-result");
const newCourseResult = document.getElementById("new-course-result");
const newEnrollmentResult = document.getElementById("new-enrollment-result");
const newCourseSessionTip = document.getElementById("new-course-session-tip");
const newEnrollmentSessionTip = document.getElementById("new-enrollment-session-tip");

const sessionEditModal = document.getElementById("session-edit-modal");
const sessionEditResult = document.getElementById("session-edit-result");

let currentSessionId = null;

async function handleResponse(resp) {
  const contentType = resp.headers.get("content-type") || "";
  if (!contentType.includes("application/json")) {
    const text = await resp.text();
    throw new Error(`接口返回非JSON（HTTP ${resp.status}）：${text.slice(0, 160)}`);
  }
  const payload = await resp.json();
  if (!resp.ok || !payload.ok) {
    throw new Error(payload.error || `请求失败（HTTP ${resp.status}）`);
  }
  return payload.data;
}

function showResult(el, html, isError = false) {
  el.classList.remove("hidden");
  el.innerHTML = isError ? `<p class="error">${html}</p>` : html;
}

function toIsoLocal(value) {
  return value ? `${value}T00:00:00` : "";
}

function isoToDateValue(value) {
  return (value || "").slice(0, 10);
}

function resetAddTrainingModal() {
  currentSessionId = null;
  document.getElementById("new-title").value = "";
  document.getElementById("new-start-date").value = "";
  document.getElementById("new-end-date").value = "";
  document.getElementById("new-location").value = "";
  document.getElementById("new-training-goal").value = "";
  document.getElementById("new-notice").value = "";
  document.getElementById("baidu-access-token").value = "";
  document.getElementById("new-course-word").value = "";
  document.getElementById("new-enrollment-file").value = "";
  newSessionResult.classList.add("hidden");
  newCourseResult.classList.add("hidden");
  newEnrollmentResult.classList.add("hidden");
  newCourseSessionTip.textContent = "当前未绑定培训班，请先完成第①步。";
  newEnrollmentSessionTip.textContent = "当前未绑定培训班，请先完成第①步。";
}

function openAddTrainingModal() {
  resetAddTrainingModal();
  addTrainingModal.classList.add("show");
}

function closeAddTrainingModal() {
  addTrainingModal.classList.remove("show");
}

function openSessionEditModal() {
  sessionEditModal.classList.add("show");
}

function closeSessionEditModal() {
  sessionEditModal.classList.remove("show");
  sessionEditResult.innerHTML = "";
}

async function fetchStats() {
  const year = document.getElementById("year").value.trim();
  if (!/^\d{4}$/.test(year)) {
    statsResult.innerHTML = '<p class="error">请输入四位年份。</p>';
    return;
  }

  try {
    const data = await handleResponse(await fetch(`/api/stats/year?year=${encodeURIComponent(year)}`));
    const rows = (data.top5 || []).map((item) => `<tr><td>${item.phone_norm}</td><td>${item.name || ""}</td><td>${item.count}</td></tr>`).join("");
    statsResult.innerHTML = `
      <p>年度：${year}</p>
      <p>参训人次：<strong>${data.total_enrollments}</strong>；参训人数：<strong>${data.total_people}</strong>；复训人数：<strong>${data.repeat_people}</strong></p>
      <table>
        <thead><tr><th>手机号</th><th>姓名</th><th>次数</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="3">暂无数据</td></tr>'}</tbody>
      </table>
    `;
  } catch (error) {
    statsResult.innerHTML = `<p class="error">统计失败：${error.message}</p>`;
  }
}

async function fetchHistory() {
  try {
    const rows = await handleResponse(await fetch("/api/session/history"));
    if (!rows.length) {
      historyList.innerHTML = "<p>暂无历史培训班。</p>";
      return;
    }
    historyList.innerHTML = rows.map((item) => `
      <div class="history-item">
        <div><strong>#${item.session_id} ${item.title || "未命名培训班"}</strong></div>
        <div>日期：${item.start_date || ""} ~ ${item.end_date || ""}</div>
        <div>地点：${item.location_text || ""}</div>
        <div>培训目标：${item.training_goal || ""}</div>
        <div>报名人次：${item.enrollment_count || 0}</div>
        <button data-action="edit-session" data-session-id="${item.session_id}">修改</button>
      </div>
    `).join("");
  } catch (error) {
    historyList.innerHTML = `<p class="error">加载历史培训班失败：${error.message}</p>`;
  }
}

async function createSession() {
  const title = document.getElementById("new-title").value.trim();
  if (!title) {
    showResult(newSessionResult, "请填写培训标题。", true);
    return;
  }

  const formData = new FormData();
  formData.append("title", title);
  formData.append("start_date", document.getElementById("new-start-date").value.trim());
  formData.append("end_date", document.getElementById("new-end-date").value.trim());
  formData.append("location_text", document.getElementById("new-location").value.trim());
  formData.append("training_goal", document.getElementById("new-training-goal").value.trim());
  const notice = document.getElementById("new-notice").files[0];
  if (notice) formData.append("notice_file", notice);

  try {
    const data = await handleResponse(await fetch("/api/session/create", { method: "POST", body: formData }));
    currentSessionId = data.session_id;
    showResult(newSessionResult, `<p>培训班创建成功，session_id：<strong>${currentSessionId}</strong></p>`);
    newCourseSessionTip.textContent = `已绑定培训班 session_id：${currentSessionId}`;
    newEnrollmentSessionTip.textContent = `已绑定培训班 session_id：${currentSessionId}`;
    await fetchHistory();
  } catch (error) {
    showResult(newSessionResult, `创建失败：${error.message}`, true);
  }
}

async function importCourseWord() {
  if (!currentSessionId) {
    showResult(newCourseResult, "请先完成第①步创建培训班。", true);
    return;
  }
  const file = document.getElementById("new-course-word").files[0];
  if (!file) {
    showResult(newCourseResult, "请先选择 Word 课程表文件。", true);
    return;
  }

  const formData = new FormData();
  formData.append("session_id", String(currentSessionId));
  formData.append("word_file", file);

  try {
    const data = await handleResponse(await fetch("/api/course/import", { method: "POST", body: formData }));
    showResult(newCourseResult, `<p>课程导入成功：新增 ${data.imported_courses} 条。</p>`);
  } catch (error) {
    showResult(newCourseResult, `课程导入失败：${error.message}`, true);
  }
}

async function importEnrollment() {
  if (!currentSessionId) {
    showResult(newEnrollmentResult, "请先完成第①步创建培训班。", true);
    return;
  }
  const file = document.getElementById("new-enrollment-file").files[0];
  if (!file) {
    showResult(newEnrollmentResult, "请先选择报名 Excel 文件。", true);
    return;
  }

  const formData = new FormData();
  formData.append("session_id", String(currentSessionId));
  formData.append("excel_file", file);

  try {
    const data = await handleResponse(await fetch("/api/enrollment/import", { method: "POST", body: formData }));
    const errors = (data.invalid_rows || []).map((it) => `<li>sheet:${it.sheet} 行:${it.row} 原因:${it.reason}</li>`).join("");
    showResult(newEnrollmentResult, `
      <p>导入完成：sheet数 ${data.sheet_count}，有效行 ${data.valid_rows}，新增学员 ${data.new_person_count}，新增报名 ${data.new_enrollment_count}。</p>
      <ul>${errors || "<li>无异常行</li>"}</ul>
    `);
    await fetchHistory();
  } catch (error) {
    showResult(newEnrollmentResult, `人员导入失败：${error.message}`, true);
  }
}

async function editSession(sessionId) {
  try {
    const data = await handleResponse(await fetch(`/api/session/${sessionId}`));
    document.getElementById("session-edit-id").value = String(data.session_id);
    document.getElementById("session-edit-title").value = data.title || "";
    document.getElementById("session-edit-start").value = isoToDateValue(data.start_date);
    document.getElementById("session-edit-end").value = isoToDateValue(data.end_date);
    document.getElementById("session-edit-location").value = data.location_text || "";
    document.getElementById("session-edit-training-goal").value = data.training_goal || "";
    document.getElementById("session-edit-notice").value = "";
    document.getElementById("session-edit-course-word").value = "";
    document.getElementById("session-edit-enrollment-file").value = "";
    sessionEditResult.innerHTML = "";
    openSessionEditModal();
  } catch (error) {
    alert(`读取培训班失败：${error.message}`);
  }
}

async function saveSessionEdit() {
  const sessionId = document.getElementById("session-edit-id").value.trim();
  const title = document.getElementById("session-edit-title").value.trim();
  if (!sessionId || !title) {
    sessionEditResult.innerHTML = '<p class="error">请填写完整信息（至少 session_id 与标题）。</p>';
    return;
  }

  const formData = new FormData();
  formData.append("session_id", sessionId);
  formData.append("title", title);
  formData.append("start_date", document.getElementById("session-edit-start").value.trim());
  formData.append("end_date", document.getElementById("session-edit-end").value.trim());
  formData.append("location_text", document.getElementById("session-edit-location").value.trim());
  formData.append("training_goal", document.getElementById("session-edit-training-goal").value.trim());
  const notice = document.getElementById("session-edit-notice").files[0];
  if (notice) formData.append("notice_file", notice);

  try {
    await handleResponse(await fetch("/api/session/update", { method: "POST", body: formData }));
    sessionEditResult.innerHTML = "<p>培训班信息保存成功。</p>";
    await fetchHistory();
  } catch (error) {
    sessionEditResult.innerHTML = `<p class="error">保存失败：${error.message}</p>`;
  }
}

async function parseNoticeAndFill() {
  const notice = document.getElementById("new-notice").files[0];
  if (!notice) {
    showResult(newSessionResult, "请先选择通知文件，再进行解析。", true);
    return;
  }
  const accessToken = document.getElementById("baidu-access-token").value.trim();
  if (!accessToken) {
    showResult(newSessionResult, "请先输入百度云 Access Token。", true);
    return;
  }

  const formData = new FormData();
  formData.append("notice_file", notice);
  formData.append("baidu_access_token", accessToken);

  try {
    showResult(newSessionResult, "正在解析通知文件，请稍候...");
    const data = await handleResponse(await fetch("/api/session/parse_notice", { method: "POST", body: formData }));
    if (data.title) document.getElementById("new-title").value = data.title;
    if (data.start_date) document.getElementById("new-start-date").value = data.start_date;
    if (data.end_date) document.getElementById("new-end-date").value = data.end_date;
    if (data.location_text) document.getElementById("new-location").value = data.location_text;
    if (data.training_goal) document.getElementById("new-training-goal").value = data.training_goal;
    showResult(
      newSessionResult,
      `<p>解析成功，已自动填充字段。</p>
       <p>培训标题：${data.title || ""}</p>
       <p>培训时间：${data.start_date || ""}${data.end_date ? ` ~ ${data.end_date}` : ""}</p>
       <p>培训目标：${data.training_goal || ""}</p>`
    );
  } catch (error) {
    showResult(newSessionResult, `通知解析失败：${error.message}`, true);
  }
}

async function reimportSessionCourse() {
  const sessionId = document.getElementById("session-edit-id").value.trim();
  if (!sessionId) {
    sessionEditResult.innerHTML = '<p class="error">请先选择一个历史培训班。</p>';
    return;
  }
  const file = document.getElementById("session-edit-course-word").files[0];
  if (!file) {
    sessionEditResult.innerHTML = '<p class="error">请先选择 Word 课程表文件。</p>';
    return;
  }

  const formData = new FormData();
  formData.append("session_id", sessionId);
  formData.append("word_file", file);
  try {
    const data = await handleResponse(await fetch("/api/course/import", { method: "POST", body: formData }));
    sessionEditResult.innerHTML = `<p>课程表重新导入成功：新增 ${data.imported_courses} 条课程。</p>`;
    document.getElementById("session-edit-course-word").value = "";
  } catch (error) {
    sessionEditResult.innerHTML = `<p class="error">课程表导入失败：${error.message}</p>`;
  }
}

async function reimportSessionEnrollment() {
  const sessionId = document.getElementById("session-edit-id").value.trim();
  if (!sessionId) {
    sessionEditResult.innerHTML = '<p class="error">请先选择一个历史培训班。</p>';
    return;
  }
  const file = document.getElementById("session-edit-enrollment-file").files[0];
  if (!file) {
    sessionEditResult.innerHTML = '<p class="error">请先选择报名 Excel 文件。</p>';
    return;
  }

  const formData = new FormData();
  formData.append("session_id", sessionId);
  formData.append("excel_file", file);
  try {
    const data = await handleResponse(await fetch("/api/enrollment/import", { method: "POST", body: formData }));
    const errors = (data.invalid_rows || []).map((it) => `<li>sheet:${it.sheet} 行:${it.row} 原因:${it.reason}</li>`).join("");
    sessionEditResult.innerHTML = `
      <p>学员名单重新导入完成：sheet数 ${data.sheet_count}，有效行 ${data.valid_rows}，新增学员 ${data.new_person_count}，新增报名 ${data.new_enrollment_count}。</p>
      <ul>${errors || "<li>无异常行</li>"}</ul>
    `;
    document.getElementById("session-edit-enrollment-file").value = "";
    await fetchHistory();
  } catch (error) {
    sessionEditResult.innerHTML = `<p class="error">学员名单导入失败：${error.message}</p>`;
  }
}

async function generateTodayTasks() {
  try {
    const data = await handleResponse(await fetch("/api/tasks/generate_today", { method: "POST" }));
    todayTaskResult.innerHTML = `<p>任务生成完成：新增 ${data.generated} 条，跳过 ${data.skipped} 条。</p>`;
    await fetchTodayTasks();
  } catch (error) {
    todayTaskResult.innerHTML = `<p class="error">生成任务失败：${error.message}</p>`;
  }
}

async function fetchTodayTasks() {
  try {
    const tasks = await handleResponse(await fetch("/api/tasks/today"));
    if (!tasks.length) {
      todayTaskList.innerHTML = "<p>今天暂无课后待发送任务。</p>";
      return;
    }

    todayTaskList.innerHTML = tasks.map((item) => {
      const submitted = Number(item.survey_submitted_count || 0);
      const total = Number(item.enrollment_total_count || 0);
      const ratioText = total > 0 ? `${((submitted / total) * 100).toFixed(1)}%` : "--";
      const surveyHtml = item.survey_link
        ? `<div>问卷：<a href="${item.survey_link}" target="_blank">${item.survey_link}</a></div>
           <div style="margin-top:6px;padding:8px;border:1px solid #e7e7e7;border-radius:6px;background:#fff;">
             <strong>问卷快速结果</strong>
             <div>填写份数：${submitted}</div>
             <div>总人数：${total}</div>
             <div>回收比例：${ratioText}</div>
           </div>`
        : "";
      const qrHtml = item.qr_data_uri
        ? `<div style="margin-top:6px;"><div>二维码：</div><img src="${item.qr_data_uri}" alt="问卷二维码" width="120" /></div>`
        : '<div class="inline-tip" style="margin-top:6px;">二维码暂不可用（请检查 Pillow/qrcode 依赖或日志）。</div>';
      const statusText = item.status === "sent" ? "已发送" : "待发送";
      const safeContent = (item.content || "").replace(/"/g, "&quot;");
      return `
        <div class="task-item">
          <div><strong>${item.course_title || "课程"}</strong>（课后）</div>
          <div>计划发送：${item.planned_at}</div>
          <div>内容：${item.content || ""}</div>
          <div>状态：${statusText}</div>
          ${surveyHtml}
          ${qrHtml}
          <div>
            <button data-action="copy-task" data-task-id="${item.task_id}" data-content="${safeContent}">复制文案</button>
            <button data-action="mark-sent" data-task-id="${item.task_id}">标记已发送</button>
          </div>
        </div>`;
    }).join("");
  } catch (error) {
    todayTaskList.innerHTML = `<p class="error">加载今天任务失败：${error.message}</p>`;
  }
}

async function fetchLogs() {
  try {
    const data = await handleResponse(await fetch("/api/logs/recent?lines=200"));
    logPath.textContent = `日志文件：${data.log_path}`;
    logContent.textContent = (data.lines || []).join("\n") || "暂无日志";
  } catch (error) {
    logPath.textContent = `日志加载失败：${error.message}`;
    logContent.textContent = "";
  }
}

async function markTaskSent(taskId) {
  try {
    await handleResponse(await fetch(`/api/tasks/${taskId}/mark_sent`, { method: "POST" }));
    await fetchTodayTasks();
  } catch (error) {
    alert(`标记失败：${error.message}`);
  }
}

async function copyTaskContent(content) {
  try {
    await navigator.clipboard.writeText(content || "");
    alert("文案已复制");
  } catch (error) {
    alert(`复制失败：${error.message}`);
  }
}

function bindEvents() {
  document.getElementById("fetch-stats").addEventListener("click", fetchStats);
  document.getElementById("export-year").addEventListener("click", () => {
    const year = document.getElementById("year").value.trim();
    if (!/^\d{4}$/.test(year)) {
      statsResult.innerHTML = '<p class="error">请输入四位年份再导出。</p>';
      return;
    }
    window.location.href = `/api/export/year?year=${encodeURIComponent(year)}`;
  });

  document.getElementById("open-add-training").addEventListener("click", openAddTrainingModal);
  document.getElementById("close-add-training").addEventListener("click", closeAddTrainingModal);
  document.getElementById("create-session").addEventListener("click", createSession);
  document.getElementById("parse-notice-btn").addEventListener("click", parseNoticeAndFill);
  document.getElementById("import-course-word").addEventListener("click", importCourseWord);
  document.getElementById("import-enrollment").addEventListener("click", importEnrollment);

  document.getElementById("refresh-history").addEventListener("click", fetchHistory);

  document.getElementById("close-session-edit").addEventListener("click", closeSessionEditModal);
  document.getElementById("save-session-edit").addEventListener("click", saveSessionEdit);
  document.getElementById("reimport-session-course").addEventListener("click", reimportSessionCourse);
  document.getElementById("reimport-session-enrollment").addEventListener("click", reimportSessionEnrollment);

  document.getElementById("generate-today-tasks").addEventListener("click", generateTodayTasks);
  document.getElementById("refresh-today-tasks").addEventListener("click", fetchTodayTasks);
  document.getElementById("refresh-logs").addEventListener("click", fetchLogs);

  historyList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action='edit-session']");
    if (!button) return;
    editSession(Number(button.dataset.sessionId));
  });

  todayTaskList.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) return;
    const action = button.dataset.action;
    if (action === "mark-sent") {
      await markTaskSent(Number(button.dataset.taskId));
    }
    if (action === "copy-task") {
      await copyTaskContent(button.dataset.content || "");
    }
  });
}

function init() {
  const currentYear = new Date().getFullYear();
  document.getElementById("year").value = String(currentYear);
  bindEvents();
  fetchStats();
  fetchHistory();
  fetchTodayTasks();
  fetchLogs();
}

init();
