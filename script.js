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
  // Total per month. Rows already have a "month" field now.
  const monthly = d3.rollups(
    data,
    rows => d3.sum(rows, d => d.amount),
    d => d.month
  )
    .map(([month, total]) => ({ month, total }))
    .sort((a, b) => d3.ascending(a.month, b.month));

  console.log(data);     // every parsed row
  console.log(monthly);  // the 12 monthly totals

  const width = 800;
  const height = 400;
  const margin = { top: 20, right: 20, bottom: 40, left: 70 };

  const svg = d3.select("#chart")
    .append("svg")
    .attr("width", width)
    .attr("height", height);

  const x = d3.scaleBand()
    .domain(monthly.map(d => d.month))
    .range([margin.left, width - margin.right])
    .padding(0.2);

  const y = d3.scaleLinear()
    .domain([0, d3.max(monthly, d => d.total)])
    .nice()
    .range([height - margin.bottom, margin.top]);

  svg.selectAll("rect")
    .data(monthly)
    .join("rect")
    .attr("class", "bar")
    .attr("x", d => x(d.month))
    .attr("y", d => y(d.total))
    .attr("width", x.bandwidth())
    .attr("height", d => height - margin.bottom - y(d.total));

  const labelOf = m => d3.timeFormat("%b %y")(new Date(m + "-01T00:00"));

  svg.append("g")
    .attr("transform", `translate(0,${height - margin.bottom})`)
    .call(d3.axisBottom(x).tickFormat(labelOf));

  svg.append("g")
    .attr("transform", `translate(${margin.left},0)`)
    .call(d3.axisLeft(y).tickFormat(d3.format("$,.0f")));
}