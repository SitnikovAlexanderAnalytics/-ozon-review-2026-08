const state = { data: null, animal: "all", method: "all", ptype: "all", week: "all", cardsShown: 20 };
const CARDS_PAGE_STEP = 20;

const fmtDate = (iso) => {
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
};

// ISO-8601 неделя года: считаем по четвергу той же недели, как в стандарте.
function isoWeek(iso) {
  const d = new Date(iso + "T00:00:00Z");
  const day = (d.getUTCDay() + 6) % 7; // понедельник = 0
  d.setUTCDate(d.getUTCDate() - day + 3); // четверг этой недели
  const firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  const firstDay = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDay + 3);
  const week = 1 + Math.round((d - firstThursday) / (7 * 86400000));
  return `${d.getUTCFullYear()}-W${String(week).padStart(2, "0")}`;
}

async function main() {
  const res = await fetch("data.json?v=13");
  const data = await res.json();
  state.data = data;

  document.getElementById("updated-date").textContent = fmtDate(data.generated_at);

  renderKpis(data);
  renderChart("chart", data.analytics);
  renderChartTable("chart-table", data.analytics);
  renderConclusions(data.conclusions);
  renderTaxonomy(data.taxonomy);
  renderWeekFilterOptions();
  renderCards();
  wireFilters();
  wireTableToggle("toggle-table", "chart", "chart-table-wrap");

  renderChart("protein-chart", data.protein_analytics);
  renderChartTable("protein-table", data.protein_analytics);
  wireTableToggle("toggle-table-protein", "protein-chart", "protein-table-wrap");
  document.getElementById("protein-note").textContent =
    `База 100% = ${data.protein_base} SKU, где протеин назван на лицевой стороне упаковки, из ${data.protein_total_skus} SKU всего.`;

  renderChart("ingredient-chart", data.ingredient_analytics);
  renderChartTable("ingredient-table", data.ingredient_analytics);
  wireTableToggle("toggle-table-ingredient", "ingredient-chart", "ingredient-table-wrap");
  document.getElementById("ingredient-note").textContent =
    `База 100% = вся выборка (${data.ingredient_base}).`;

  renderChart("producer-chart", data.producer_analytics);
  renderChartTable("producer-table", data.producer_analytics);
  wireTableToggle("toggle-table-producer", "producer-chart", "producer-table-wrap");
  document.getElementById("producer-note").textContent =
    `База 100% = ${data.producer_base} товаров, где производитель однозначно определён ` +
    `(для РФ — по названию юрлица, иначе считается неизвестным), из ${data.n_products} в подробно разобранной выборке.`;
}

function renderKpis(data) {
  const items = [
    { value: data.n_products, label: "Новых SKU разобрано" },
    { value: data.days_tracked, label: "Дней наблюдений" },
  ];
  const el = document.getElementById("kpis");
  el.innerHTML = items.map(i => `
    <div class="kpi">
      <div class="value">${i.value}</div>
      <div class="label">${i.label}</div>
    </div>
  `).join("");
}

function renderChart(elId, analytics) {
  const max = Math.max(...analytics.map(a => a.count));
  const el = document.getElementById(elId);
  el.innerHTML = analytics.map(a => `
    <div class="bar-row">
      <div class="bar-label">${a.name}</div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${(a.count / max * 100).toFixed(1)}%"></div>
      </div>
      <div class="bar-value">${a.count} (${a.pct}%)</div>
    </div>
  `).join("");
}

function renderChartTable(tableId, analytics) {
  const tbody = document.querySelector(`#${tableId} tbody`);
  tbody.innerHTML = analytics.map(a => `
    <tr><td>${a.name}</td><td>${a.count}</td><td>${a.pct}%</td></tr>
  `).join("");
}

function wireTableToggle(btnId, chartId, tableWrapId) {
  const btn = document.getElementById(btnId);
  const chart = document.getElementById(chartId);
  const tableWrap = document.getElementById(tableWrapId);
  let showingTable = false;
  btn.addEventListener("click", () => {
    showingTable = !showingTable;
    tableWrap.classList.toggle("hidden", !showingTable);
    chart.classList.toggle("hidden", showingTable);
    btn.setAttribute("aria-pressed", String(showingTable));
    btn.textContent = showingTable ? "Показать диаграммой" : "Показать таблицей";
  });
}

function renderConclusions(conclusions) {
  const el = document.getElementById("conclusions");
  el.innerHTML = conclusions.map(c => `<li>${escapeHtml(c).replace(/^\d+\.\s*/, "")}</li>`).join("");
}

function renderTaxonomy(taxonomy) {
  const tbody = document.querySelector("#taxonomy-table tbody");
  tbody.innerHTML = taxonomy.map(t => `
    <tr>
      <td><strong>${t.name}</strong></td>
      <td>${t.description}</td>
      <td>${t.examples}</td>
    </tr>
  `).join("");
}

function renderCards() {
  const { products } = state.data;
  const filtered = products
    .filter(p =>
      (state.animal === "all" || p.animal === state.animal) &&
      (state.method === "all" || p.method === state.method) &&
      (state.ptype === "all" || p.ptype === state.ptype) &&
      (state.week === "all" || isoWeek(p.date) === state.week)
    )
    .sort((a, b) => b.date.localeCompare(a.date) || b.oid.localeCompare(a.oid));

  const el = document.getElementById("cards");
  if (filtered.length === 0) {
    el.innerHTML = `<p style="color:var(--text-muted)">Нет товаров по выбранным фильтрам.</p>`;
    return;
  }

  const visible = filtered.slice(0, state.cardsShown);
  const remaining = filtered.length - visible.length;
  const moreHtml = remaining > 0
    ? `<button class="btn-ghost" id="cards-more" type="button">Показать ещё ${Math.min(CARDS_PAGE_STEP, remaining)} (осталось ${remaining})</button>`
    : "";

  el.innerHTML = visible.map(cardHtml).join("") + moreHtml;

  const moreBtn = document.getElementById("cards-more");
  if (moreBtn) {
    moreBtn.addEventListener("click", () => {
      state.cardsShown += CARDS_PAGE_STEP;
      renderCards();
    });
  }
}

function cardHtml(p) {
  return `
    <article class="card">
      <div class="card-photo"><img src="${p.photo}" alt="Упаковка: ${escapeHtml(p.name)}"></div>
      <div class="card-body">
        <div class="card-badges">
          <span class="badge ${p.method === 'diff' ? 'badge-diff' : ''}">${p.method === 'diff' ? 'Подтверждено ID' : 'Кандидат («Новинки»)'}</span>
          <span class="badge">${fmtDate(p.date)}</span>
        </div>
        <div class="card-brand">${escapeHtml(p.brand)} · ${escapeHtml(p.animal)} · ${escapeHtml(p.ptype)}</div>
        <p class="card-name">${escapeHtml(p.name)}</p>
        <div class="card-tags">
          ${p.claims.length
            ? p.claims.map(c => `<span class="tag" title="${escapeHtml(c.text)}">${escapeHtml(c.category)}</span>`).join("")
            : `<span class="tag">Без заявленных клеймов</span>`}
        </div>
        <dl class="card-meta">
          <div><dt>Протеин</dt><dd>${escapeHtml(p.protein || "—")}</dd></div>
          <div><dt>Ингредиенты (лицевая сторона упаковки)</dt><dd>${escapeHtml(p.ingredients || "—")}</dd></div>
          <div><dt>Завод-изготовитель</dt><dd>${escapeHtml(p.plant || "—")}</dd></div>
        </dl>
        <a class="card-link" href="${p.url}" target="_blank" rel="noopener">Открыть на Ozon →</a>
      </div>
    </article>
  `;
}

function wireFilters() {
  document.querySelectorAll("[data-filter-animal]").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll("[data-filter-animal]").forEach(b => b.classList.remove("is-active"));
      btn.classList.add("is-active");
      state.animal = btn.dataset.filterAnimal;
      state.cardsShown = 20;
      renderCards();
    });
  });
  document.querySelectorAll("[data-filter-method]").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll("[data-filter-method]").forEach(b => b.classList.remove("is-active"));
      btn.classList.add("is-active");
      state.method = btn.dataset.filterMethod;
      state.cardsShown = 20;
      renderCards();
    });
  });
  document.querySelectorAll("[data-filter-ptype]").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll("[data-filter-ptype]").forEach(b => b.classList.remove("is-active"));
      btn.classList.add("is-active");
      state.ptype = btn.dataset.filterPtype;
      state.cardsShown = 20;
      renderCards();
    });
  });
}

function wireWeekFilter() {
  document.querySelectorAll("[data-filter-week]").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll("[data-filter-week]").forEach(b => b.classList.remove("is-active"));
      btn.classList.add("is-active");
      state.week = btn.dataset.filterWeek;
      state.cardsShown = 20;
      renderCards();
    });
  });
}

function renderWeekFilterOptions() {
  const { products } = state.data;
  const weeks = [...new Set(products.map(p => isoWeek(p.date)))].sort();
  const group = document.getElementById("week-filter-group");
  weeks.forEach(w => {
    const btn = document.createElement("button");
    btn.className = "chip";
    btn.type = "button";
    btn.dataset.filterWeek = w;
    btn.textContent = w.replace(/^\d+-W/, "Неделя ");
    group.appendChild(btn);
  });
  wireWeekFilter();
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[c]));
}

main();
