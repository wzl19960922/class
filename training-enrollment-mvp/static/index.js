const statsResult = document.getElementById("stats-result");
const historyList = document.getElementById("history-list");
const courseResult = document.getElementById("course-result");
const courseList = document.getElementById("course-list");
const todayTaskResult = document.getElementById("today-task-result");
const todayTaskList = document.getElementById("today-task-list");
const courseEditModal = document.getElementById("course-edit-modal");
const courseListModal = document.getElementById("course-list-modal");
const sessionEditModal = document.getElementById("session-edit-modal");

function renderTop5(items) {
  if (!items || items.length === 0) return "<p>无</p>";
  const rows = items.map((item) => `<tr><td>${item.phone_norm}</td><td>${item.name || ""}</td><td>${item.count}</td></tr>`);
  return `<table><thead><tr><th>手机号</th><th>姓名</th><th>次数</th></tr></thead><tbody>${rows.join("")}</tbody></table>`;
}

async function handleResponse(response) {
  const contentType = response.headers.get("content-type") || "";
  if (!contentType.includes("application/json")) {
    const text = await response.text();
    throw new Error(`接口未返回JSON（HTTP ${response.status}）: ${text.slice(0, 120)}`);
  }
  const payload = await response.json();
  if (!payload.ok) throw new Error(payload.error || "请求失败");
  return payload.data;
}

function resetCourseForm() {
  document.getElementById("course-title").value = "";
  document.getElementById("course-teacher").value = "";
  document.getElementById("course-start-at").value = "";
  document.getElementById("course-end-at").value = "";
  document.getElementById("course-location").value = "";
}

function closeCourseEditModal() {
  courseEditModal.classList.remove("show");
}

function openCourseEditModal() {
  courseEditModal.classList.add("show");
}

function openCourseListModal() {
  courseListModal.classList.add("show");
}

function closeCourseListModal() {
  courseListModal.classList.remove("show");
}

function closeSessionEditModal() {
  sessionEditModal.classList.remove("show");
}

function openSessionEditModal() {
  sessionEditModal.classList.add("show");
}

async function fetchStats() {
  const year = document.getElementById("year").value.trim();
  if (!year) return (statsResult.innerHTML = '<p class="error">请输入年份。</p>');
  try {
    const data = await handleResponse(await fetch(`/api/stats/year?year=${encodeURIComponent(year)}`));
    statsResult.innerHTML = `<p>参训人次: ${data.total_enrollments}</p><p>参训人数: ${data.total_people}</p><p>复训人数: ${data.repeat_people}</p><p>Top5:</p>${renderTop5(data.top5)}`;
  } catch (error) {
    statsResult.innerHTML = `<p class="error">${error.message}</p>`;
  }
}

async function fetchHistory() {
  try {
    const sessions = await handleResponse(await fetch("/api/session/history"));
    if (!sessions?.length) return (historyList.innerHTML = "<p>暂无历史培训记录。</p>");
    historyList.innerHTML = sessions.map((item) => `<div class="history-item"><div><strong>#${item.session_id} ${item.title || "未命名培训"}</strong></div><div>时间：${item.start_date || ""} ~ ${item.end_date || ""}</div><div>地点：${item.location_text || ""}</div><div>报名人次：${item.enrollment_count}</div><div>创建时间：${item.created_at || ""}</div><div><button onclick="editSession(${item.session_id})">弹窗修改</button></div></div>`).join("");
  } catch (error) {
    historyList.innerHTML = `<p class="error">加载历史培训失败：${error.message}</p>`;
  }
}


async function editSession(sessionId) {
  try {
    const data = await handleResponse(await fetch(`/api/session/${sessionId}`));
    document.getElementById("session-edit-id").value = data.session_id;
    document.getElementById("session-edit-title").value = data.title || "";
    document.getElementById("session-edit-start-date").value = data.start_date || "";
    document.getElementById("session-edit-end-date").value = data.end_date || "";
    document.getElementById("session-edit-location").value = data.location_text || "";
    openSessionEditModal();
  } catch (error) {
    historyList.innerHTML = `<p class="error">加载历史培训失败：${error.message}</p>`;
  }
}

async function saveSessionEdit() {
  const sessionId = document.getElementById("session-edit-id").value;
  if (!sessionId) return;

  const formData = new FormData();
  formData.append("session_id", sessionId);
  formData.append("title", document.getElementById("session-edit-title").value.trim());
  formData.append("start_date", document.getElementById("session-edit-start-date").value.trim());
  formData.append("end_date", document.getElementById("session-edit-end-date").value.trim());
  formData.append("location_text", document.getElementById("session-edit-location").value.trim());
  const noticeFile = document.getElementById("session-edit-notice").files[0];
  if (noticeFile) {
    formData.append("notice_file", noticeFile);
  }

  try {
    await handleResponse(await fetch("/api/session/update", { method: "POST", body: formData }));
    closeSessionEditModal();
    await fetchHistory();
  } catch (error) {
    historyList.innerHTML = `<p class="error">保存历史培训失败：${error.message}</p>`;
  }
}

async function importCourseTable() {
  const wordFile = document.getElementById("course-word-file").files[0];
  if (!wordFile) return (courseResult.innerHTML = '<p class="error">请先选择 Word 文件。</p>');
  if (!wordFile.name.toLowerCase().endsWith(".docx")) return (courseResult.innerHTML = '<p class="error">仅支持 .docx 文件。</p>');

  const sessionId = document.getElementById("course-session-id").value.trim();
  if (sessionId && !/^\d+$/.test(sessionId)) return (courseResult.innerHTML = '<p class="error">session_id 只能是数字。</p>');

  const formData = new FormData();
  formData.append("word_file", wordFile);
  if (sessionId) formData.append("session_id", sessionId);

  try {
    const data = await handleResponse(await fetch("/api/course/import", { method: "POST", body: formData }));
    courseResult.innerHTML = `<p>课程导入成功：${data.imported_courses} 条。</p>`;
    await fetchCourses();
  } catch (error) {
    courseResult.innerHTML = `<p class="error">课程导入失败：${error.message}</p>`;
  }
}

async function saveCourse() {
  const payload = {
    title: document.getElementById("course-title").value.trim(),
    teacher: document.getElementById("course-teacher").value.trim(),
    start_at: document.getElementById("course-start-at").value.trim(),
    end_at: document.getElementById("course-end-at").value.trim(),
    location: document.getElementById("course-location").value.trim(),
    session_id: document.getElementById("course-session-id").value.trim(),
  };
  if (!payload.title) return (courseResult.innerHTML = '<p class="error">课程名称不能为空。</p>');

  try {
    const data = await handleResponse(await fetch("/api/course/create", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) }));
    courseResult.innerHTML = `<p>课程保存成功，course_id: ${data.course_id}</p>`;
    resetCourseForm();
    await fetchCourses();
  } catch (error) {
    courseResult.innerHTML = `<p class="error">课程保存失败：${error.message}</p>`;
  }
}

async function editCourse(courseId) {
  try {
    const data = await handleResponse(await fetch(`/api/course/${courseId}`));
    document.getElementById("course-edit-id").value = data.course_id;
    document.getElementById("course-edit-title").value = data.title || "";
    document.getElementById("course-edit-teacher").value = data.teacher || "";
    document.getElementById("course-edit-start-at").value = data.start_at || "";
    document.getElementById("course-edit-end-at").value = data.end_at || "";
    document.getElementById("course-edit-location").value = data.location || "";
    document.getElementById("course-edit-session-id").value = data.session_id || "";
    openCourseEditModal();
  } catch (error) {
    courseResult.innerHTML = `<p class="error">加载课程失败：${error.message}</p>`;
  }
}

async function saveCourseEdit() {
  const payload = {
    course_id: Number(document.getElementById("course-edit-id").value),
    title: document.getElementById("course-edit-title").value.trim(),
    teacher: document.getElementById("course-edit-teacher").value.trim(),
    start_at: document.getElementById("course-edit-start-at").value.trim(),
    end_at: document.getElementById("course-edit-end-at").value.trim(),
    location: document.getElementById("course-edit-location").value.trim(),
    session_id: document.getElementById("course-edit-session-id").value.trim(),
  };
  if (!payload.course_id || !payload.title) return alert("课程ID或课程名称不能为空");

  try {
    await handleResponse(await fetch("/api/course/update", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) }));
    closeCourseEditModal();
    courseResult.innerHTML = `<p>课程 ${payload.course_id} 更新成功。</p>`;
    await fetchCourses();
  } catch (error) {
    courseResult.innerHTML = `<p class="error">更新课程失败：${error.message}</p>`;
  }
}

async function deleteCourse(courseId) {
  if (!confirm(`确认删除课程 ${courseId} 吗？`)) return;
  try {
    await handleResponse(await fetch(`/api/course/${courseId}/delete`, { method: "POST" }));
    courseResult.innerHTML = `<p>课程 ${courseId} 已删除。</p>`;
    await fetchCourses();
  } catch (error) {
    courseResult.innerHTML = `<p class="error">删除课程失败：${error.message}</p>`;
  }
}

async function fetchCourses() {
  try {
    const courses = await handleResponse(await fetch("/api/course/list"));
    if (!courses?.length) return (courseList.innerHTML = "<p>暂无课程记录。</p>");
    courseList.innerHTML = courses.map((item) => `<div class="course-item"><div><strong>${item.title}</strong></div><div>老师：${item.teacher || ""}</div><div>开始：${item.start_at || ""}</div><div>结束：${item.end_at || ""}</div><div>地点：${item.location || ""}</div><div>session_id：${item.session_id || "未绑定"}</div><div><button onclick="editCourse(${item.course_id})">编辑</button><button onclick="deleteCourse(${item.course_id})">删除</button></div></div>`).join("");
  } catch (error) {
    courseList.innerHTML = `<p class="error">加载课程失败：${error.message}</p>`;
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

async function markTaskSent(taskId) {
  try {
    await handleResponse(await fetch(`/api/tasks/${taskId}/mark_sent`, { method: "POST" }));
    await fetchTodayTasks();
  } catch (error) {
    alert(`标记失败：${error.message}`);
  }
}

async function copyTaskContent(taskId, content) {
  try {
    await navigator.clipboard.writeText(content);
    alert(`任务 ${taskId} 文案已复制`);
  } catch {
    alert("复制失败，请手动复制。");
  }
}

async function fetchTodayTasks() {
  try {
    const tasks = await handleResponse(await fetch("/api/tasks/today"));
    if (!tasks?.length) return (todayTaskList.innerHTML = "<p>今天暂无待发送任务。</p>");
    todayTaskList.innerHTML = tasks.map((item) => {
      const qrHtml = item.qr_data_uri ? `<div><img src="${item.qr_data_uri}" alt="二维码" width="120" /></div>` : "";
      const surveyHtml = item.survey_link ? `<div>问卷：<a href="${item.survey_link}" target="_blank">${item.survey_link}</a></div>` : "";
      const statusText = item.status === "sent" ? "已发送" : "待发送";
      const escapedContent = (item.content || "").replace(/"/g, "&quot;");
      return `<div class="task-item"><div><strong>${item.course_title || "课程"}</strong>（${item.task_type === "pre" ? "课前" : "课后"}）</div><div>计划发送：${item.planned_at}</div><div>内容：${item.content || ""}</div><div>状态：${statusText}</div>${surveyHtml}${qrHtml}<div><button onclick="copyTaskContent(${item.task_id}, \"${escapedContent}\")">复制文案</button><button onclick="markTaskSent(${item.task_id})">标记已发送</button></div></div>`;
    }).join("");
  } catch (error) {
    todayTaskList.innerHTML = `<p class="error">加载今天任务失败：${error.message}</p>`;
  }
}

document.getElementById("fetch-stats").addEventListener("click", fetchStats);
document.getElementById("export-year").addEventListener("click", () => {
  const year = document.getElementById("year").value.trim();
  if (!year) return (statsResult.innerHTML = '<p class="error">请输入年份再导出。</p>');
  window.location.href = `/api/export/year?year=${encodeURIComponent(year)}`;
});
document.getElementById("refresh-history").addEventListener("click", fetchHistory);
document.getElementById("import-course").addEventListener("click", importCourseTable);
document.getElementById("save-course").addEventListener("click", saveCourse);
document.getElementById("reset-course-form").addEventListener("click", resetCourseForm);
document.getElementById("open-course-list").addEventListener("click", async () => {
  await fetchCourses();
  openCourseListModal();
});
document.getElementById("refresh-courses").addEventListener("click", async () => {
  await fetchCourses();
  openCourseListModal();
});
document.getElementById("close-course-list").addEventListener("click", closeCourseListModal);
document.getElementById("generate-today-tasks").addEventListener("click", generateTodayTasks);
document.getElementById("refresh-today-tasks").addEventListener("click", fetchTodayTasks);
document.getElementById("save-course-edit").addEventListener("click", saveCourseEdit);
document.getElementById("close-course-edit").addEventListener("click", closeCourseEditModal);
document.getElementById("save-session-edit").addEventListener("click", saveSessionEdit);
document.getElementById("close-session-edit").addEventListener("click", closeSessionEditModal);

window.copyTaskContent = copyTaskContent;
window.markTaskSent = markTaskSent;
window.editCourse = editCourse;
window.deleteCourse = deleteCourse;
window.editSession = editSession;

const currentYear = new Date().getFullYear();
document.getElementById("year").value = String(currentYear);
fetchStats();
fetchHistory();
fetchCourses();
fetchTodayTasks();
