const YEAR = 2026;

// ---------- Step 1: load and parse the workbook ----------
async function loadData() {
  const buffer = await fetch("data/2026.xlsx").then(r => {
    if (!r.ok) throw new Error("Could not find data/2026.xlsx");
    return r.arrayBuffer();
  });
  const wb = XLSX.read(buffer, { cellDates: true });
  const { rows, warnings } = parseWorkbook(wb, YEAR);
  warnings.forEach(w => console.warn(w));
  return rows;
}

loadData()
  .then(drawChart)
  .catch(err => console.error("Something went wrong:", err));

// ---------- Step 2: draw the chart ----------
function drawChart(data) {

  // ---------- Step 1: shape the data ----------
  // Group by month, then by category, summing amounts at the bottom level.
  const totals = d3.rollup(
    data,
    rows => d3.sum(rows, d => d.amount),
    d => d.month,
    d => d.category
  );

  const categories = [...new Set(data.map(d => d.category))];
  const months = [...totals.keys()].sort();

  // One row per month: { month: "2026-01", "Housing & Utilities": 2063, Food: 599, ... }
  const table = months.map(month => {
    const row = { month };
    categories.forEach(c => row[c] = totals.get(month).get(c) || 0);
    return row;
  });

  console.log(table);

  // d3.stack works out where each segment starts and ends.
  // Result: one array per category, with a [bottom, top] pair for each month.
  const series = d3.stack().keys(categories)(table);

  // ---------- Step 2: canvas and scales ----------
  const width = 900;
  const height = 420;
  const margin = { top: 30, right: 180, bottom: 40, left: 70 };

  const svg = d3.select("#chart")
    .append("svg")
    .attr("width", width)
    .attr("height", height);

  const x = d3.scaleBand()
    .domain(months)
    .range([margin.left, width - margin.right])
    .padding(0.2);

  const y = d3.scaleLinear()
    .domain([0, d3.max(series, s => d3.max(s, d => d[1]))])  // tallest stack
    .nice()
    .range([height - margin.bottom, margin.top]);

  const color = d3.scaleOrdinal()
    .domain(categories)
    .range(d3.schemeTableau10);

  // ---------- Step 3: draw the stacked bars ----------
  const tooltip = d3.select("#tooltip");
  const money = d3.format("$,.2f");
  const pct = d3.format(".0%");
  const monthName = m => d3.timeFormat("%B %Y")(new Date(m + "-01T00:00"));
  const monthTotal = row => d3.sum(categories, c => row[c]);

  svg.append("g")
    .selectAll("g")
    .data(series)
    .join("g")
    .attr("fill", s => color(s.key))
    .selectAll("rect")
    .data(s => s.map(seg => ({ key: s.key, seg })))
    .join("rect")
    .attr("class", "segment")
    .attr("x", d => x(d.seg.data.month))
    .attr("y", d => y(d.seg[1]))
    .attr("width", x.bandwidth())
    .attr("height", d => y(d.seg[0]) - y(d.seg[1]))
    .on("mouseenter", function () {
      d3.select(this).attr("stroke", "black").attr("stroke-width", 2);
    })
    .on("mousemove", (event, d) => {
      const amount = d.seg[1] - d.seg[0];
      const total = monthTotal(d.seg.data);
      tooltip
        .style("opacity", 1)
        .html(
          `<strong>${d.key}</strong> <br>` +
          `${money(amount)} (${pct(amount / total)} of the month)`
        )
        .style("left", event.pageX + 14 + "px")
        .style("top", event.pageY - 10 + "px");
    })
    .on("mouseleave", function () {
      d3.select(this).attr("stroke", null);
      tooltip.style("opacity", 0);
    });

  // Total label above each bar
  svg.append("g")
    .selectAll("text")
    .data(table)
    .join("text")
    .attr("x", d => x(d.month) + x.bandwidth() / 2)
    .attr("y", d => y(monthTotal(d)) - 6)
    .attr("text-anchor", "middle")
    .style("font-size", "11px")
    .style("font-weight", "bold")
    .text(d => d3.format("$,.0f")(monthTotal(d)));

  // ---------- Step 4: axes ----------
  const labelOf = m => d3.timeFormat("%b %y")(new Date(m + "-01T00:00"));

  svg.append("g")
    .attr("transform", `translate(0,${height - margin.bottom})`)
    .call(d3.axisBottom(x).tickFormat(labelOf));

  svg.append("g")
    .attr("transform", `translate(${margin.left},0)`)
    .call(d3.axisLeft(y).tickFormat(d3.format("$,.0f")));

  // ---------- Step 5: legend ----------
  const legend = svg.append("g")
    .attr("transform", `translate(${width - margin.right + 20},${margin.top})`);

  [...categories].reverse().forEach((c, i) => {
    const row = legend.append("g").attr("transform", `translate(0,${i * 24})`);
    row.append("rect").attr("width", 14).attr("height", 14).attr("fill", color(c));
    row.append("text").attr("x", 22).attr("y", 12).style("font-size", "13px").text(c);
  });
}