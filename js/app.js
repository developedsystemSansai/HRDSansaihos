// =====================================
// CONFIG
// =====================================
const CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbwtnfYjlEAloUtZt5cvblr_6hUAgaVx4hYzff86RmkjjDRgGFbIM_43jnJnFkIn4eq8eQ/exec",
  TARGET: 15,
  REFRESH_INTERVAL: 30000
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
// 🔥 JSONP FETCH (แก้ CORS)
// =====================================
function fetchJSONP(url) {
  return new Promise((resolve, reject) => {

    const callbackName = "jsonp_" + Date.now();

    window[callbackName] = function (data) {
      resolve(data);
      document.body.removeChild(script);
      delete window[callbackName];
    };

    const script = document.createElement("script");
    script.src = url + "&callback=" + callbackName;

    script.onerror = () => reject("โหลด API ไม่ได้");

    document.body.appendChild(script);
  });
}

// =====================================
// LOAD DASHBOARD
// =====================================
async function loadDashboard() {

  try {

    const url = CONFIG.API_URL + "?action=dashboard";

    const res = await fetchJSONP(url);

    state.data = res.data || [];
    state.summary = res.summary || {};
    state.dept = res.dept || {};

    renderAll();

  } catch (err) {
    console.error(err);
    alert("โหลดข้อมูลไม่ได้");
  }
}

// =====================================
// RENDER
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
  total.innerText = state.summary.total || 0;
  pass.innerText = state.summary.pass || 0;
  fail.innerText = state.summary.fail || 0;
}

// =====================================
// TABLE
// =====================================
function renderTable(data) {

  let html = "";

  data.forEach(d => {

    let percent = Math.min((d.hours / CONFIG.TARGET) * 100, 100);

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
      <td style="color:${d.status==='ผ่าน'?'green':'red'}">
        ${d.status}
      </td>
    </tr>`;
  });

  table.innerHTML = html;
}
function initSearch() {

  const search = document.getElementById("search");

  // 🔥 เช็คก่อนใช้
  if (!search) {
    console.warn("❌ ไม่พบ search element");
    return;
  }

  search.addEventListener("input", function () {

    const keyword = this.value.toLowerCase();

    const filtered = state.data.filter(d =>
      d.name.toLowerCase().includes(keyword)
    );

    renderTable(filtered);
  });
}

// =====================================
// CHART
// =====================================
function renderChart() {

  let labels = Object.keys(state.dept);
  let values = Object.values(state.dept);

  if (state.chart) {
    state.chart.data.labels = labels;
    state.chart.data.datasets[0].data = values;
    state.chart.update();
    return;
  }

  state.chart = new Chart(chart, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'ชั่วโมงรวม',
        data: values
      }]
    }
  });
}

// =====================================
// DETAIL
// =====================================
async function openDetail(name) {

  modal.style.display = "block";
  modalName.innerText = name;
  modalTable.innerHTML = "Loading...";

  try {

    const url = CONFIG.API_URL + "?action=detail&name=" + encodeURIComponent(name);

    const data = await fetchJSONP(url);

    let html = "";

    data.forEach(d => {
      html += `
      <tr>
        <td>${d.course}</td>
        <td>${d.hours}</td>
        <td>${d.date}</td>
      </tr>`;
    });

    modalTable.innerHTML = html;

  } catch {
    modalTable.innerHTML = "โหลดไม่ได้";
  }
}

function closeModal() {
  modal.style.display = "none";
}

// =====================================
// REALTIME
// =====================================
function startRealtime() {
  setInterval(loadDashboard, CONFIG.REFRESH_INTERVAL);
}

// =====================================
// START
// =====================================
document.addEventListener("DOMContentLoaded", () => {
  loadDashboard();
  startRealtime();
});
