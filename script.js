const YEAR = 2026;
const TAGS_KEY = "finance-tags-v2";      // { rowId: { category, sub } }
const OLD_TAGS_KEY = "finance-tags-v1";  // earlier version: { rowId: "Category" }
const TREE_KEY = "finance-tree-v1";      // the super category > category > subcategory tree
const NEW_ITEM = "__new__";
const SEP = " › ";

// Starting point for the tree. Everything here can be changed on the Categories tab.
const DEFAULT_GROUPS = ["Necessity", "Non-necessity", "Saving"];
const DEFAULT_GROUP_OF = {
  "Housing & Utilities": "Necessity",
  "Transportation": "Necessity",
  "Food": "Necessity",
  "Personal": "Non-necessity"
};

// ---------- State ----------
let rawRows = [];                  // parsed from Excel. Never modified.
const rawById = new Map();         // rowId -> raw row
let tags = {};                     // your tags: { rowId: { category, sub } }  (sub is null for "no subcategory")
let tree = { groups: [], categories: [] };   // categories: [{ name, group, subs: [] }]
let selectedPie = null;            // "2026-03", "all", or null
let detail = null;                 // what the drawer shows: { month: "2026-03" | "all", key: string | null }
let lastDetailSig = "";            // used to reset the drawer's scroll when the selection changes
let applySame = false;             // "row tags apply to every row with the same name"
const chartView = { level: "category", scope: "all", onlySubs: true };

// ---------- Shared helpers ----------
const money = d3.format("$,.2f");
const pct = d3.format(".0%");
const labelOf = m => d3.timeFormat("%b %y")(new Date(m + "-01T00:00"));
const monthName = m => d3.timeFormat("%B %Y")(new Date(m + "-01T00:00"));
const tooltip = d3.select("#tooltip");
const color = d3.scaleOrdinal([...d3.schemeTableau10, ...d3.schemeSet3]);   // grows as new names appear

function loadJSON(key, fallback) {
  try { const v = JSON.parse(localStorage.getItem(key)); return v ?? fallback; }
  catch { return fallback; }
}
function save() {
  try {
    localStorage.setItem(TAGS_KEY, JSON.stringify(tags));
    localStorage.setItem(TREE_KEY, JSON.stringify(tree));
  } catch (e) { console.warn("Could not save:", e); }
}

// ---------- The tree ----------
const findCat = name => tree.categories.find(c => c.name === name);

function ensureGroup(g) {
  if (!tree.groups.includes(g)) tree.groups.push(g);
}
function ensureCategory(name) {
  let c = findCat(name);
  if (!c) {
    c = { name, group: DEFAULT_GROUP_OF[name] || "Unsorted", subs: [] };
    ensureGroup(c.group);
    tree.categories.push(c);
  }
  return c;
}
function ensureSub(catName, sub) {
  const c = ensureCategory(catName);
  if (sub && !c.subs.includes(sub)) c.subs.push(sub);
}

const originalSub = r => (r.subcategory === r.category ? null : r.subcategory);

// Make sure every category/subcategory in use exists in the tree
function syncTree() {
  tree.categories.forEach(c => ensureGroup(c.group));
  rawRows.forEach(r => {
    const t = tags[r.id];
    ensureSub(t ? t.category : r.category, t ? t.sub : originalSub(r));
  });
}

// A row as the charts see it: spreadsheet values, replaced by your tag if there is one
function effective(row) {
  const t = tags[row.id];
  const category = t ? t.category : row.category;
  const sub = t ? t.sub : originalSub(row);
  const c = findCat(category);
  return { ...row, category, sub, group: c ? c.group : "Unsorted" };
}

// Categories in display order: by super category, then the order they were created
const catsInOrder = () => tree.groups.flatMap(g => tree.categories.filter(c => c.group === g));

// Names for the chart's stack segments at each level
function keyOf(r, level) {
  if (level === "group") return r.group;
  if (level === "category") return r.category;
  return r.sub ? r.category + SEP + r.sub : r.category;
}
function pathText(r, level) {
  if (level === "group") return r.group;
  if (level === "category") return [r.group, r.category].join(SEP);
  return [r.group, r.category, r.sub].filter(Boolean).join(SEP);
}
function orderedKeys(level) {
  if (level === "group") return [...tree.groups];
  const cats = catsInOrder();
  if (level === "category") return cats.map(c => c.name);
  return cats.flatMap(c => [c.name, ...c.subs.map(s => c.name + SEP + s)]);
}

function inScope(r) {
  if (chartView.scope === "all") return true;
  const i = chartView.scope.indexOf(":");
  const kind = chartView.scope.slice(0, i);
  const name = chartView.scope.slice(i + 1);
  if (kind === "g") return r.group === name;
  if (kind === "c") return r.category === name;
  const [cat, sub] = JSON.parse(name);              // kind === "s": one subcategory
  return r.category === cat && r.sub === sub;
}

// The rows the charts are built from (your tags, scope and subcategory filter applied)
function chartRows() {
  const subsOnly = chartView.level === "sub" && chartView.onlySubs;
  return rawRows.map(effective).filter(inScope).filter(r => !subsOnly || r.sub);
}

// The rows behind whatever was clicked on the charts
function detailRows() {
  const level = chartView.level;
  return chartRows()
    .filter(r => detail.month === "all" || r.month === detail.month)
    .filter(r => detail.key === null || keyOf(r, level) === detail.key)
    .sort((a, b) => d3.ascending(a.month, b.month) || d3.descending(a.amount, b.amount));
}

// ---------- Load ----------
async function loadData() {
  const buffer = await fetch("data/2026.xlsx").then(r => {
    if (!r.ok) throw new Error("Could not find data/2026.xlsx");
    return r.arrayBuffer();
  });
  const wb = XLSX.read(buffer, { cellDates: true });
  const { rows, warnings } = parseWorkbook(wb, YEAR);
  warnings.forEach(w => console.warn(w));

  // Stable ids so tags can be remembered. Duplicates get #1, #2.
  const seen = {};
  rows.forEach(r => {
    const base = `${r.month}|${r.item.trim()}|${r.amount.toFixed(2)}`;
    seen[base] = (seen[base] || 0) + 1;
    r.id = `${base}|${seen[base]}`;
  });
  return rows;
}

loadData()
  .then(rows => {
    rawRows = rows;
    rows.forEach(r => rawById.set(r.id, r));

    tree = loadJSON(TREE_KEY, null) || { groups: [...DEFAULT_GROUPS], categories: [] };
    tags = loadJSON(TAGS_KEY, null);
    if (!tags) {
      // First run of this version: convert tags saved by the earlier version
      tags = {};
      Object.entries(loadJSON(OLD_TAGS_KEY, {})).forEach(([id, cat]) => { tags[id] = { category: cat, sub: null }; });
    }
    syncTree();
    save();

    setupTabs();
    setupChartControls();
    setupTableControls();
    d3.select("#add-group").on("click", addGroup);
    showView("charts");
  })
  .catch(err => console.error("Something went wrong:", err));

// =====================================================
//  Tabs
// =====================================================
function setupTabs() {
  d3.selectAll(".tab").on("click", function () { showView(this.dataset.view); });
}

function showView(name) {
  d3.selectAll(".tab").classed("active", function () { return this.dataset.view === name; });
  ["charts", "table", "categories"].forEach(n => d3.select("#view-" + n).classed("hidden", n !== name));
  if (name === "charts") renderCharts();
  if (name === "table") renderTable();
  if (name === "categories") renderCategories();
}

// =====================================================
//  The "new category" popup (works from any screen)
// =====================================================
// Resolves with [category, sub] once saved, or null if cancelled.
function askCategory({ title = "New category", button = "Create & tag" } = {}) {
  return new Promise(resolve => {
    const modal = d3.select("#cat-dialog");
    const $cat = d3.select("#dlg-category");
    const $sub = d3.select("#dlg-sub");
    const $group = d3.select("#dlg-group");
    const $newGroup = d3.select("#dlg-newgroup");
    const $hint = d3.select("#dlg-hint");

    d3.select("#dlg-title").text(title);
    d3.select("#dlg-save").text(button);
    $cat.property("value", "");
    $sub.property("value", "");
    $newGroup.property("value", "");
    $hint.classed("error", false).text("");

    d3.select("#dlg-cat-list").selectAll("option")
      .data(tree.categories.map(c => c.name)).join("option").attr("value", d => d);

    $group.selectAll("*").remove();
    $group.append("option").attr("value", "").text("Choose a super category…");
    $group.selectAll("option.g").data(tree.groups).join("option")
      .attr("class", "g").attr("value", d => d).text(d => d);
    $group.append("option").attr("value", NEW_ITEM).text("+ New super category…");
    $group.property("value", "").property("disabled", false);

    const existingCat = () => {
      const name = $cat.property("value").trim().toLowerCase();
      return tree.categories.find(c => c.name.toLowerCase() === name);
    };

    // Keep the form in step with what's typed
    const sync = () => {
      const existing = existingCat();
      const name = $cat.property("value").trim();
      d3.select("#dlg-sub-list").selectAll("option")
        .data(existing ? existing.subs : []).join("option").attr("value", d => d);
      if (existing) $group.property("value", existing.group);
      $group.property("disabled", !!existing);
      $newGroup.classed("hidden", !!existing || $group.property("value") !== NEW_ITEM);
      $hint.classed("error", false).text(
        existing ? `"${existing.name}" already exists in ${existing.group}.`
          : name ? "New category: choose the super category it belongs to."
            : ""
      );
    };
    sync();

    const fail = msg => $hint.classed("error", true).text(msg);

    const close = result => {
      modal.classed("hidden", true);
      d3.select(document).on("keydown.dlg", null);
      resolve(result);
    };

    const submit = () => {
      const category = $cat.property("value").trim();
      const subRaw = $sub.property("value").trim();
      if (!category) return fail("Type a category name.");

      let c = existingCat();
      if (!c) {
        let group = $group.property("value");
        if (group === NEW_ITEM) group = $newGroup.property("value").trim();
        if (!group) return fail("Choose a super category for the new category.");
        ensureGroup(group);
        c = { name: category, group, subs: [] };
        tree.categories.push(c);
      }

      let sub = null;
      if (subRaw) {
        sub = c.subs.find(s => s.toLowerCase() === subRaw.toLowerCase()) ?? subRaw;
        if (!c.subs.includes(sub)) c.subs.push(sub);
      }
      save();
      close([c.name, sub]);
    };

    $cat.on("input", sync);
    $group.on("change", sync);
    [$cat, $sub, $newGroup].forEach(inp => inp.on("keydown", e => { if (e.key === "Enter") submit(); }));
    d3.select("#dlg-save").on("click", submit);
    d3.select("#dlg-cancel").on("click", () => close(null));
    modal.on("click", e => { if (e.target === modal.node()) close(null); });
    d3.select(document).on("keydown.dlg", e => { if (e.key === "Escape") close(null); });

    modal.classed("hidden", false);
    $cat.node().focus();
  });
}

// =====================================================
//  Tagging pieces shared by the Transactions tab and the chart drawer
// =====================================================
// The dropdown choices, grouped by super category: "Food", "Food › Groceries", ...
function tagOptions() {
  return tree.groups.map(g => ({
    group: g,
    items: tree.categories.filter(c => c.group === g).flatMap(c => [
      { label: c.name, value: JSON.stringify([c.name, null]) },
      ...c.subs.map(s => ({ label: c.name + SEP + s, value: JSON.stringify([c.name, s]) }))
    ])
  })).filter(o => o.items.length);
}

function fillTagSelect(select, withPlaceholder) {
  if (withPlaceholder) select.append("option").attr("value", "").text("Choose a category…");
  select.selectAll("optgroup").data(tagOptions()).join("optgroup")
    .attr("label", d => d.group)
    .selectAll("option").data(d => d.items).join("option")
    .attr("value", d => d.value).text(d => d.label);
  select.append("option").attr("value", NEW_ITEM).text("+ New category or subcategory…");
}

function setTag(rowId, category, sub) {
  const raw = rawById.get(rowId);
  if (category === raw.category && sub === originalSub(raw)) delete tags[rowId];   // same as the spreadsheet
  else tags[rowId] = { category, sub };
}

// A row's dropdown changed. onChange re-renders whichever screen we're on.
async function onTag(row, selectEl, onChange) {
  let target;
  if (selectEl.value === NEW_ITEM) {
    target = await askCategory();
    if (!target) { selectEl.value = JSON.stringify([row.category, row.sub]); return; }   // cancelled
  } else {
    target = JSON.parse(selectEl.value);
  }
  const [category, sub] = target;

  const norm = s => s.trim().toLowerCase();
  const targets = applySame ? rawRows.filter(r => norm(r.item) === norm(row.item)) : [rawById.get(row.id)];

  targets.forEach(r => setTag(r.id, category, sub));
  save();
  onChange();
}

// A table of transactions, each with a category dropdown
function buildTxTable(host, rows, onChange, { compact = false } = {}) {
  const table = host.append("table").attr("class", "tx-table");
  const cols = compact
    ? ["Month", "Item", "Amount", "Category"]
    : ["Month", "Item", "Amount", "Original", "Super category", "Your category"];
  table.append("thead").append("tr").selectAll("th").data(cols).join("th")
    .attr("class", c => (c === "Amount" ? "num" : null)).text(c => c);

  const tr = table.append("tbody").selectAll("tr").data(rows).join("tr")
    .classed("changed", d => d.id in tags);

  tr.append("td").text(d => labelOf(d.month));
  tr.append("td").text(d => d.item);
  tr.append("td").attr("class", "num").text(d => money(d.amount));
  if (!compact) {
    tr.append("td").attr("class", "orig")
      .text(d => originalSub(d) ? d.category + SEP + d.subcategory : d.category);
    tr.append("td").text(d => d.group);
  }

  // d.category / d.sub are the *effective* values (your tag, if any)
  const select = tr.append("td").append("select")
    .on("change", function (event, d) { onTag(d, this, onChange); });
  fillTagSelect(select, false);
  select.property("value", d => JSON.stringify([d.category, d.sub]));
}

// "Tag every row shown below as: [category] [button]"
function buildBulkBar(host, getRows, onChange) {
  const bar = host.append("div").attr("class", "filters bulk");
  bar.append("strong").text("Bulk:");
  const select = bar.append("select");
  fillTagSelect(select, true);

  bar.append("button").text("Tag all shown rows").on("click", async () => {
    const val = select.property("value");
    if (!val) return;
    const rows = getRows();
    if (!rows.length) return;

    const target = val === NEW_ITEM ? await askCategory() : JSON.parse(val);
    if (!target) return;
    const [category, sub] = target;
    const label = sub ? category + SEP + sub : category;
    if (!confirm(`Tag ${rows.length} rows as "${label}"?`)) { onChange(); return; }

    rows.forEach(r => setTag(r.id, category, sub));
    save();
    onChange();
  });
}

// =====================================================
//  Screen 3: edit the category tree
// =====================================================
function usageCounts() {
  const u = {};   // { category: { total, subs: { sub: n } } }
  rawRows.map(effective).forEach(r => {
    const e = (u[r.category] ??= { total: 0, subs: {} });
    e.total++;
    if (r.sub) e.subs[r.sub] = (e.subs[r.sub] || 0) + 1;
  });
  return u;
}

function askName(message) {
  return (prompt(message) || "").trim();
}

function addGroup() {
  const name = askName("Name for the new super category (e.g. Saving):");
  if (!name) return;
  if (tree.groups.includes(name)) { alert("That super category already exists."); return; }
  tree.groups.push(name);
  treeChanged();
}
function addCategory(group) {
  const name = askName(`New category in "${group}":`);
  if (!name) return;
  if (findCat(name)) { alert("A category with that name already exists."); return; }
  ensureCategory(name).group = group;
  treeChanged();
}
function addSub(catName) {
  const name = askName(`New subcategory under "${catName}":`);
  if (!name) return;
  const c = findCat(catName);
  if (c.subs.includes(name)) { alert("That subcategory already exists."); return; }
  c.subs.push(name);
  treeChanged();
}
function moveCategory(name, group) {
  findCat(name).group = group;
  treeChanged();
}
function deleteGroup(g) {
  tree.groups = tree.groups.filter(x => x !== g);
  treeChanged();
}
function deleteCategory(name) {
  tree.categories = tree.categories.filter(c => c.name !== name);
  treeChanged();
}
function deleteSub(catName, sub) {
  const c = findCat(catName);
  c.subs = c.subs.filter(s => s !== sub);
  treeChanged();
}
function treeChanged() {
  save();
  renderCategories();
}

function renderCategories() {
  syncTree();
  const usage = usageCounts();
  const root = d3.select("#tree");
  root.selectAll("*").remove();

  tree.groups.forEach(g => {
    const cats = tree.categories.filter(c => c.group === g);
    const itemCount = d3.sum(cats, c => usage[c.name]?.total || 0);

    const card = root.append("div").attr("class", "group-card");
    const head = card.append("div").attr("class", "group-head");
    head.append("h3").text(g);
    head.append("span").attr("class", "count").text(`${cats.length} categories · ${itemCount} items`);
    head.append("span").attr("class", "spacer");
    head.append("button").text("+ Category").on("click", () => addCategory(g));
    head.append("button").text("Delete")
      .property("disabled", cats.length > 0)
      .attr("title", cats.length ? "Move or delete its categories first" : null)
      .on("click", () => deleteGroup(g));

    cats.forEach(c => {
      const n = usage[c.name]?.total || 0;
      const line = card.append("div").attr("class", "cat-line");
      line.append("strong").text(c.name);
      line.append("span").attr("class", "count").text(`${n} items`);
      line.append("span").attr("class", "spacer");

      const move = line.append("select").attr("title", "Move to another super category")
        .on("change", function () { moveCategory(c.name, this.value); });
      move.selectAll("option").data(tree.groups).join("option").attr("value", d => d).text(d => d);
      move.property("value", g);

      line.append("button").text("+ Sub").on("click", () => addSub(c.name));
      line.append("button").text("Delete")
        .property("disabled", n > 0)
        .attr("title", n > 0 ? "Retag its items first" : null)
        .on("click", () => deleteCategory(c.name));

      c.subs.forEach(s => {
        const sn = usage[c.name]?.subs[s] || 0;
        const sl = card.append("div").attr("class", "sub-line");
        sl.append("span").text("↳ " + s);
        sl.append("span").attr("class", "count").text(`${sn} items`);
        sl.append("span").attr("class", "spacer");
        sl.append("button").text("Delete")
          .property("disabled", sn > 0)
          .attr("title", sn > 0 ? "Retag its items first" : null)
          .on("click", () => deleteSub(c.name, s));
      });
    });
  });
}

// =====================================================
//  Screen 2: the transactions table
// =====================================================
function setupTableControls() {
  const months = [...new Set(rawRows.map(r => r.month))].sort();
  d3.select("#filter-month").selectAll("option.m")
    .data(months).join("option").attr("class", "m")
    .attr("value", d => d).text(monthName);

  d3.select("#filter-month").on("change", renderTable);
  d3.select("#filter-search").on("input", renderTable);
  d3.select("#filter-changed").on("change", renderTable);
  d3.select("#apply-all").on("change", function () { applySame = this.checked; });

  d3.select("#new-cat-btn").on("click", async () => {
    const target = await askCategory({ title: "New category", button: "Create" });
    if (target) renderTable();   // refresh the dropdowns
  });

  d3.select("#reset-tags").on("click", () => {
    if (!confirm("Remove all item tags and go back to the spreadsheet's categories?\n(Your categories and super categories are kept.)")) return;
    tags = {};
    save();
    renderTable();
  });
}

// Rows that pass the current filters
function visibleRows() {
  const month = d3.select("#filter-month").property("value");
  const search = d3.select("#filter-search").property("value").trim().toLowerCase();
  const changedOnly = d3.select("#filter-changed").property("checked");

  return rawRows.map(effective).filter(r =>
    (month === "all" || r.month === month) &&
    (!search || [r.item, r.category, r.sub || "", r.group].join(" ").toLowerCase().includes(search)) &&
    (!changedOnly || r.id in tags)
  );
}

function renderTable() {
  syncTree();
  const rows = visibleRows();

  d3.select("#apply-all").property("checked", applySame);
  d3.select("#row-count").text(`${rows.length} of ${rawRows.length} rows · ${Object.keys(tags).length} tagged`);

  const bulkHost = d3.select("#bulk-host");
  bulkHost.selectAll("*").remove();
  buildBulkBar(bulkHost, visibleRows, renderTable);

  const host = d3.select("#tx-host");
  host.selectAll("*").remove();
  buildTxTable(host, rows, renderTable);
}

// =====================================================
//  Screen 1: stacked bars + pie
// =====================================================
function setupChartControls() {
  // Changing what the chart shows makes the old selection meaningless, so clear it
  d3.select("#level-select").on("change", function () { chartView.level = this.value; detail = null; renderCharts(); });
  d3.select("#scope-select").on("change", function () { chartView.scope = this.value; detail = null; renderCharts(); });
  d3.select("#only-subs").on("change", function () { chartView.onlySubs = this.checked; detail = null; renderCharts(); });
}

function fillScopeSelect() {
  const sel = d3.select("#scope-select");
  sel.selectAll("*").remove();
  sel.append("option").attr("value", "all").text("All spending");
  sel.append("optgroup").attr("label", "Super categories")
    .selectAll("option").data(tree.groups).join("option").attr("value", g => "g:" + g).text(g => g);
  sel.append("optgroup").attr("label", "Categories")
    .selectAll("option").data(catsInOrder().map(c => c.name)).join("option").attr("value", c => "c:" + c).text(c => c);

  const subScopes = catsInOrder().flatMap(c =>
    c.subs.map(s => ({ value: "s:" + JSON.stringify([c.name, s]), label: c.name + SEP + s })));
  sel.append("optgroup").attr("label", "Subcategories")
    .selectAll("option").data(subScopes).join("option").attr("value", d => d.value).text(d => d.label);

  const valid = ["all", ...tree.groups.map(g => "g:" + g), ...tree.categories.map(c => "c:" + c.name),
    ...subScopes.map(d => d.value)];
  if (!valid.includes(chartView.scope)) chartView.scope = "all";
  sel.property("value", chartView.scope);
}

// The drawer under the charts: transactions behind the clicked bar / month / slice / legend item
function renderDetail() {
  const panel = d3.select("#detail-panel");
  if (!detail) { panel.classed("hidden", true); return; }

  const sig = detail.month + "|" + detail.key;
  const prevScroll = sig === lastDetailSig ? (panel.select(".tx-scroll").node()?.scrollTop ?? 0) : 0;
  lastDetailSig = sig;

  const rows = detailRows();
  const total = d3.sum(rows, r => r.amount);

  panel.classed("hidden", false);
  panel.selectAll("*").remove();

  const head = panel.append("div").attr("class", "detail-head");
  head.append("h3").text([detail.month === "all" ? `Whole year ${YEAR}` : monthName(detail.month), detail.key]
    .filter(Boolean).join(" · "));
  head.append("span").attr("class", "count").text(`${rows.length} items · ${money(total)}`);
  head.append("span").attr("class", "spacer");
  head.append("button").text("+ New category").on("click", async () => {
    const target = await askCategory({ title: "New category", button: "Create" });
    if (target) renderCharts();
  });
  head.append("button").text("Close ✕").on("click", () => { detail = null; renderCharts(); });

  buildBulkBar(panel, detailRows, renderCharts);

  const same = panel.append("label").attr("class", "apply-same");
  same.append("input").attr("type", "checkbox").property("checked", applySame)
    .on("change", function () { applySame = this.checked; });
  same.append("span").text(" Row tags apply to every row with the same name");

  const scroll = panel.append("div").attr("class", "tx-scroll");
  if (!rows.length) {
    scroll.append("p").attr("class", "hint").text("Nothing left here. Everything was retagged, or this selection is empty.");
  } else {
    buildTxTable(scroll, rows, renderCharts, { compact: true });
  }
  scroll.node().scrollTop = prevScroll;
}

function renderCharts() {
  syncTree();
  tooltip.style("opacity", 0);
  fillScopeSelect();
  d3.select("#level-select").property("value", chartView.level);
  d3.select("#chart").selectAll("*").remove();

  const level = chartView.level;
  d3.select("#only-subs-label").classed("hidden", level !== "sub");
  d3.select("#only-subs").property("checked", chartView.onlySubs);

  // ---------- Step 1: shape the data (using your tags, level and scope) ----------
  const data = chartRows();

  if (!data.length) {
    d3.select("#chart").append("p").text("No spending in this selection.");
    d3.select("#pie-panel").classed("hidden", true);
    d3.select("#detail-panel").classed("hidden", true);
    return;
  }

  const key = r => keyOf(r, level);
  const info = {};                                   // key -> "Super › Category › Sub" for tooltips
  data.forEach(r => { info[key(r)] = pathText(r, level); });

  const present = new Set(data.map(key));
  const keys = orderedKeys(level).filter(k => present.has(k));
  present.forEach(k => { if (!keys.includes(k)) keys.push(k); });   // safety net
  keys.forEach(k => color(k));                                      // register colors in a stable order

  const totals = d3.rollup(data, rows => d3.sum(rows, d => d.amount), d => d.month, key);
  const months = [...new Set(rawRows.map(r => r.month))].sort();

  const table = months.map(month => {
    const row = { month };
    keys.forEach(k => row[k] = totals.get(month)?.get(k) || 0);
    return row;
  });

  // Diverging offset keeps negative amounts (refunds, paybacks) below the zero line
  const series = d3.stack().keys(keys).offset(d3.stackOffsetDiverging)(table);
  const monthTotal = row => d3.sum(keys, k => row[k]);

  // Is this segment the one shown in the drawer?
  const isSelected = d => detail && detail.key === d.key &&
    (detail.month === "all" || detail.month === d.seg.data.month);

  // ---------- Step 2: canvas and scales ----------
  const width = 960;
  const height = Math.max(420, 70 + keys.length * 20);   // taller when the legend is long
  const margin = { top: 30, right: 220, bottom: 40, left: 70 };

  const svg = d3.select("#chart").append("svg").attr("width", width).attr("height", height);

  const x = d3.scaleBand().domain(months).range([margin.left, width - margin.right]).padding(0.2);

  const yMin = Math.min(0, d3.min(series, s => d3.min(s, d => Math.min(d[0], d[1]))));
  const yMax = d3.max(series, s => d3.max(s, d => Math.max(d[0], d[1])));
  const y = d3.scaleLinear().domain([yMin, yMax]).nice().range([height - margin.bottom, margin.top]);

  // ---------- Step 3: stacked bars with hover and click ----------
  svg.append("g")
    .selectAll("g")
    .data(series)
    .join("g")
    .attr("fill", s => color(s.key))
    .selectAll("rect")
    .data(s => s.map(seg => ({ key: s.key, seg })).filter(p => p.seg[1] !== p.seg[0]))
    .join("rect")
    .attr("class", "segment")
    .attr("x", d => x(d.seg.data.month))
    .attr("y", d => y(Math.max(d.seg[0], d.seg[1])))
    .attr("width", x.bandwidth())
    .attr("height", d => Math.abs(y(d.seg[0]) - y(d.seg[1])))
    .attr("stroke", d => (isSelected(d) ? "black" : null))
    .attr("stroke-width", 2)
    .on("mouseenter", function () { d3.select(this).attr("stroke", "black"); })
    .on("mousemove", (event, d) => {
      const amount = d.seg[1] - d.seg[0];
      const total = monthTotal(d.seg.data);
      tooltip
        .style("opacity", 1)
        .html(
          `<strong>${d.key}</strong> · ${monthName(d.seg.data.month)}<br>` +
          `<span style="opacity:.7">${info[d.key]}</span><br>` +
          `${money(amount)} (${pct(amount / total)} of the month)<br>` +
          `Month total: ${money(total)}`
        )
        .style("left", event.pageX + 14 + "px")
        .style("top", event.pageY - 10 + "px");
    })
    .on("mouseleave", function (event, d) {
      d3.select(this).attr("stroke", isSelected(d) ? "black" : null);
      tooltip.style("opacity", 0);
    })
    .on("click", (event, d) => {
      selectedPie = d.seg.data.month;                       // the pie follows along
      detail = { month: d.seg.data.month, key: d.key };
      renderCharts();
    });

  // Total label above each bar
  svg.append("g")
    .selectAll("text")
    .data(table)
    .join("text")
    .attr("x", d => x(d.month) + x.bandwidth() / 2)
    .attr("y", (d, i) => y(Math.max(0, d3.max(series, s => s[i][1]))) - 6)
    .attr("text-anchor", "middle")
    .style("font-size", "11px")
    .style("font-weight", "bold")
    .text(d => d3.format("$,.0f")(monthTotal(d)));

  // ---------- Step 4: axes ----------
  const xAxis = svg.append("g")
    .attr("transform", `translate(0,${height - margin.bottom})`)
    .call(d3.axisBottom(x).tickFormat(labelOf));

  xAxis.selectAll(".tick")
    .classed("clickable-tick", true)
    .on("click", (event, month) => {
      selectedPie = month;
      detail = { month, key: null };                        // everything in that month
      renderCharts();
    });

  svg.append("g")
    .attr("transform", `translate(${margin.left},0)`)
    .call(d3.axisLeft(y).tickFormat(d3.format("$,.0f")));

  // ---------- Step 5: legend (click a name to see its transactions for the year) ----------
  const legend = svg.append("g")
    .attr("transform", `translate(${width - margin.right + 20},${margin.top})`);

  [...keys].reverse().forEach((k, i) => {
    const row = legend.append("g")
      .attr("class", "legend-item")
      .attr("transform", `translate(0,${i * 20})`)
      .on("click", () => { detail = { month: "all", key: k }; renderCharts(); });
    row.append("rect").attr("width", 12).attr("height", 12).attr("fill", color(k));
    row.append("text").attr("x", 20).attr("y", 10).style("font-size", "12px")
      .style("font-weight", detail && detail.key === k ? "bold" : null).text(k);
  });

  // ---------- Step 6: pie for a clicked month ----------
  d3.select("#year-btn").on("click", () => showPie("all"));

  function showPie(selection) {
    selectedPie = selection;

    const slices = keys.map(k => ({
      key: k,
      value: selection === "all" ? d3.sum(table, r => r[k]) : table.find(r => r.month === selection)[k]
    })).filter(d => d.value > 0);   // a pie can't show negative slices

    const total = d3.sum(slices, d => d.value);

    d3.select("#pie-panel").classed("hidden", false);
    d3.select("#pie-title").text(selection === "all" ? `Whole year ${YEAR}` : monthName(selection));

    xAxis.selectAll(".tick").each(function (m) {
      d3.select(this).select("text").style("font-weight", m === selection ? "bold" : null);
    });

    const size = 300;
    const radius = size / 2 - 10;

    const root = d3.select("#pie");
    root.selectAll("*").remove();

    const g = root.append("svg").attr("width", size).attr("height", size)
      .append("g").attr("transform", `translate(${size / 2},${size / 2})`);

    const arc = d3.arc().innerRadius(radius * 0.5).outerRadius(radius);
    const pie = d3.pie().value(d => d.value).sort(null);

    g.selectAll("path")
      .data(pie(slices))
      .join("path")
      .attr("class", "segment")
      .attr("d", arc)
      .attr("fill", d => color(d.data.key))
      .attr("stroke", "white")
      .attr("stroke-width", 2)
      .on("mousemove", (event, d) => {
        tooltip
          .style("opacity", 1)
          .html(`<strong>${d.data.key}</strong><br><span style="opacity:.7">${info[d.data.key]}</span><br>` +
            `${money(d.data.value)} (${pct(d.data.value / total)})`)
          .style("left", event.pageX + 14 + "px")
          .style("top", event.pageY - 10 + "px");
      })
      .on("mouseleave", () => tooltip.style("opacity", 0))
      .on("click", (event, d) => {
        detail = { month: selection, key: d.data.key };
        renderCharts();
      });

    g.append("text").attr("text-anchor", "middle").attr("dy", "0.35em")
      .style("font-weight", "bold").text(d3.format("$,.0f")(total));

    d3.select("#pie-list").selectAll("div").remove();
    d3.select("#pie-list").selectAll("div")
      .data(slices).join("div").attr("class", "pie-row")
      .html(d => `<span class="swatch" style="background:${color(d.key)}"></span>` +
        `${d.key}<span class="amt">${money(d.value)} (${pct(d.value / total)})</span>`);
  }

  // Keep the pie open on the same month after anything changes
  if (selectedPie) showPie(selectedPie);
  else d3.select("#pie-panel").classed("hidden", true);

  renderDetail();
}