const statsResult = document.getElementById("stats-result");
const historyList = document.getElementById("history-list");
const courseResult = document.getElementById("course-result");
const courseList = document.getElementById("course-list");
const todayTaskResult = document.getElementById("today-task-result");
const todayTaskList = document.getElementById("today-task-list");

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
            <div><a href="/add?session_id=${item.session_id}">修改信息/重新上传材料</a></div>
          </div>
        `
      )
      .join("");
  } catch (error) {
    historyList.innerHTML = `<p class="error">加载历史培训失败：${error.message}</p>`;
  }
}

async function importCourseTable() {
  const fileInput = document.getElementById("course-word-file");
  const wordFile = fileInput.files[0];
  if (!wordFile) {
    courseResult.innerHTML = '<p class="error">请先选择 Word 文件。</p>';
    return;
  }
  if (!wordFile.name.toLowerCase().endsWith(".docx")) {
    courseResult.innerHTML = '<p class="error">仅支持 .docx 文件。</p>';
    return;
  }

  const sessionId = document.getElementById("course-session-id").value.trim();
  if (sessionId && !/^\d+$/.test(sessionId)) {
    courseResult.innerHTML = '<p class="error">session_id 只能是数字。</p>';
    return;
  }

  const formData = new FormData();
  formData.append("word_file", wordFile);
  if (sessionId) {
    formData.append("session_id", sessionId);
  }

  try {
    const response = await fetch("/api/course/import", { method: "POST", body: formData });
    const data = await handleResponse(response);
    courseResult.innerHTML = `<p>课程导入成功：${data.imported_courses} 条。</p>`;
    await fetchCourses();
  } catch (error) {
    courseResult.innerHTML = `<p class="error">课程导入失败：${error.message}</p>`;
  }
}

async function fetchCourses() {
  try {
    const response = await fetch("/api/course/list");
    const courses = await handleResponse(response);
    if (!courses || courses.length === 0) {
      courseList.innerHTML = "<p>暂无课程记录。</p>";
      return;
    }

    courseList.innerHTML = courses
      .map(
        (item) => `
          <div class="course-item">
            <div><strong>${item.title}</strong></div>
            <div>老师：${item.teacher || ""}</div>
            <div>开始：${item.start_at || ""}</div>
            <div>结束：${item.end_at || ""}</div>
            <div>地点：${item.location || ""}</div>
            <div>session_id：${item.session_id || "未绑定"}</div>
          </div>
        `
      )
      .join("");
  } catch (error) {
    courseList.innerHTML = `<p class="error">加载课程失败：${error.message}</p>`;
  }
}

async function generateTodayTasks() {
  try {
    const response = await fetch("/api/tasks/generate_today", { method: "POST" });
    const data = await handleResponse(response);
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
    if (!tasks || tasks.length === 0) {
      todayTaskList.innerHTML = "<p>今天暂无待发送任务。</p>";
      return;
    }

    todayTaskList.innerHTML = tasks
      .map((item) => {
        const qrHtml = item.qr_data_uri ? `<div><img src="${item.qr_data_uri}" alt="二维码" width="120" /></div>` : "";
        const surveyHtml = item.survey_link ? `<div>问卷：<a href="${item.survey_link}" target="_blank">${item.survey_link}</a></div>` : "";
        const statusText = item.status === "sent" ? "已发送" : "待发送";
        const escapedContent = (item.content || "").replace(/"/g, "&quot;");
        return `
          <div class="task-item">
            <div><strong>${item.course_title || "课程"}</strong>（${item.task_type === "pre" ? "课前" : "课后"}）</div>
            <div>计划发送：${item.planned_at}</div>
            <div>内容：${item.content || ""}</div>
            <div>状态：${statusText}</div>
            ${surveyHtml}
            ${qrHtml}
            <div>
              <button onclick="copyTaskContent(${item.task_id}, \"${escapedContent}\")">复制文案</button>
              <button onclick="markTaskSent(${item.task_id})">标记已发送</button>
            </div>
          </div>
        `;
      })
      .join("");
  } catch (error) {
    todayTaskList.innerHTML = `<p class="error">加载今天任务失败：${error.message}</p>`;
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
document.getElementById("import-course").addEventListener("click", importCourseTable);
document.getElementById("refresh-courses").addEventListener("click", fetchCourses);
document.getElementById("generate-today-tasks").addEventListener("click", generateTodayTasks);
document.getElementById("refresh-today-tasks").addEventListener("click", fetchTodayTasks);

window.copyTaskContent = copyTaskContent;
window.markTaskSent = markTaskSent;

fetchHistory();
fetchCourses();
fetchTodayTasks();
