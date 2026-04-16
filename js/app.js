// =====================================
// 🔥 CONFIG (แก้แค่บรรทัดนี้)
// =====================================
const CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbwtnfYjlEAloUtZt5cvblr_6hUAgaVx4hYzff86RmkjjDRgGFbIM_43jnJnFkIn4eq8eQ/exec", // 
  TARGET: 15,
  REFRESH_INTERVAL: 30000 // refresh ทุก 30 วิ
};

// =====================================
// STATE
// =====================================
let state = {
  data: [],
  summary: {},
  dept: {},
  chart: null
};

// =====================================
// 🔥 SAFE FETCH (กัน error JSON พัง)
// =====================================
async function fetchJSON(url) {

  console.log("📡 Fetch:", url);

  try {
    const res = await fetch(url);

    const text = await res.text();

    // ❌ ถ้าเป็น HTML = ยิงผิด URL
    if (text.startsWith("<!DOCTYPE")) {
      throw new Error("❌ API ไม่ใช่ JSON → ตรวจสอบ URL GAS");
    }

    return JSON.parse(text);

  } catch (err) {
    console.error("❌ FETCH ERROR:", err);
    throw err;
  }
}

// =====================================
// LOAD DASHBOARD
// =====================================
async function loadDashboard() {

  try {
    showLoading(true);

    const url = CONFIG.API_URL + "?action=dashboard";

    const res = await fetchJSON(url);

    state.data = res.data || [];
    state.summary = res.summary || {};
    state.dept = res.dept || {};

    renderAll();

  } catch (err) {

    console.error("❌ LOAD ERROR:", err.message);

    alert("โหลดข้อมูลไม่ได้\n\n" + err.message);

  } finally {
    showLoading(false);
  }
}

// =====================================
// RENDER ALL
// =====================================
function renderAll() {
  renderSummary();
  renderTable(state.data);
  renderChart();
}

// =====================================
// SUMMARY
// =====================================
function renderSummary() {
  document.getElementById("total").innerText = state.summary.total || 0;
  document.getElementById("pass").innerText = state.summary.pass || 0;
  document.getElementById("fail").innerText = state.summary.fail || 0;
}

// =====================================
// TABLE
// =====================================
function renderTable(data) {

  const table = document.getElementById("table");

  if (!data.length) {
    table.innerHTML = "<tr><td colspan='4'>ไม่มีข้อมูล</td></tr>";
    return;
  }

  let html = "";

  data.forEach(d => {

    const percent = Math.min((d.hours / CONFIG.TARGET) * 100, 100);

    html += `
    <tr onclick="openDetail('${d.name}')">
      <td>${d.name}</td>
      <td>${d.dept}</td>
      <td>
        <div class="progress">
          <div class="progress-bar" style="width:${percent}%"></div>
        </div>
        ${d.hours}/${CONFIG.TARGET}
      </td>
      <td style="color:${d.status === 'ผ่าน' ? '#059669' : '#dc2626'};font-weight:600;">
        ${d.status}
      </td>
    </tr>`;
  });

  table.innerHTML = html;
}

// =====================================
// CHART
// =====================================
function renderChart() {

  const labels = Object.keys(state.dept);
  const values = Object.values(state.dept);

  const ctx = document.getElementById("chart");

  if (state.chart) {
    state.chart.data.labels = labels;
    state.chart.data.datasets[0].data = values;
    state.chart.update();
    return;
  }

  state.chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'ชั่วโมงรวม',
        data: values
      }]
    },
    options: {
      responsive: true
    }
  });
}

// =====================================
// DETAIL MODAL
// =====================================
async function openDetail(name) {

  const modal = document.getElementById("modal");
  const modalName = document.getElementById("modalName");
  const modalTable = document.getElementById("modalTable");

  modal.style.display = "block";
  modalName.innerText = name;
  modalTable.innerHTML = "<tr><td colspan='3'>กำลังโหลด...</td></tr>";

  try {

    const url = CONFIG.API_URL + "?action=detail&name=" + encodeURIComponent(name);

    const data = await fetchJSON(url);

    if (!data.length) {
      modalTable.innerHTML = "<tr><td colspan='3'>ไม่มีข้อมูล</td></tr>";
      return;
    }

    let html = "";

    data.forEach(d => {
      html += `
      <tr>
        <td>${d.course}</td>
        <td>${d.hours}</td>
        <td>${d.date || "-"}</td>
      </tr>`;
    });

    modalTable.innerHTML = html;

  } catch (err) {
    modalTable.innerHTML = "<tr><td colspan='3'>โหลดไม่ได้</td></tr>";
  }
}

// =====================================
// CLOSE MODAL
// =====================================
function closeModal() {
  document.getElementById("modal").style.display = "none";
}

// =====================================
// LOADING UI
// =====================================
function showLoading(show) {
  const loader = document.getElementById("loader");
  if (!loader) return;

  loader.style.display = show ? "flex" : "none";
}

// =====================================
// REALTIME AUTO REFRESH
// =====================================
function startRealtime() {

  setInterval(() => {
    console.log("🔄 refresh...");
    loadDashboard();
  }, CONFIG.REFRESH_INTERVAL);
}

// =====================================
// SEARCH
// =====================================
function initSearch() {

  const search = document.getElementById("search");

  if (!search) return;

  search.addEventListener("input", function () {

    const keyword = this.value.toLowerCase();

    const filtered = state.data.filter(d =>
      d.name.toLowerCase().includes(keyword)
    );

    renderTable(filtered);
  });
}

// =====================================
// START
// =====================================
document.addEventListener("DOMContentLoaded", () => {

  console.log("🚀 App Start");

  initSearch();
  loadDashboard();
  startRealtime();

});
