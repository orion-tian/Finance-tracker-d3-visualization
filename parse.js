// Turns the one-sheet-per-month budget workbook into flat rows:
// { month: "2026-01", category, subcategory, item, amount }

// Edit this if you add a category. Subcategories go in the array.
const SPENDING_STRUCTURE = {
  "Housing & Utilities": [],
  "Transportation": [],
  "Food": ["Groceries", "Eating Out", "Work Lunch"],
  "Personal": []
};

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"];

// Label column / amount column pairs (0-based): C/D, F/G, I/J
const BLOCKS = [[2, 3], [5, 6], [8, 9]];

const isNum = v => typeof v === "number" && !isNaN(v);
const cleanLabel = v => (v == null || v instanceof Date) ? null : String(v).trim();

function parseWorkbook(wb, year) {
  const out = [];
  const warnings = [];

  // Look up which category each header name belongs to
  const parentOf = {};
  Object.entries(SPENDING_STRUCTURE).forEach(([cat, subs]) => {
    parentOf[cat] = { category: cat, subcategory: cat };
    subs.forEach(s => parentOf[s] = { category: cat, subcategory: s });
  });

  wb.SheetNames.forEach(name => {
    const monthIdx = MONTH_NAMES.indexOf(name.trim());
    if (monthIdx < 0) return; // skip sheets that aren't months
    const month = `${year}-${String(monthIdx + 1).padStart(2, "0")}`;
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });

    BLOCKS.forEach(([lc, vc]) => {
      // Find header rows in this column pair
      const headers = [];
      grid.forEach((row, r) => {
        const label = cleanLabel(row[lc]);
        if (r > 0 && label && parentOf[label] && isNum(row[vc])) {
          headers.push({ r, label, total: row[vc] });
        }
      });

      headers.forEach(h => {
        const { category, subcategory } = parentOf[h.label];
        // "Food" itself is just the sum of its subcategories, so skip it
        if (h.label === category && SPENDING_STRUCTURE[category].length > 0) return;

        // Read rows beneath the header until the running sum matches its total
        let sum = 0;
        const items = [];
        for (let r = h.r + 1; r < grid.length; r++) {
          const label = cleanLabel(grid[r][lc]);
          const amt = grid[r][vc];
          if (label && parentOf[label] && isNum(amt)) break; // reached the next header
          if (!isNum(amt) || label === "Average") continue;
          const item = grid[r][lc] instanceof Date
            ? grid[r][lc].toISOString().slice(0, 10)  // Work Lunch rows are dates
            : (label || "(unlabeled)");
          items.push({ item, amount: amt });
          sum += amt;
          if (Math.abs(sum - h.total) < 0.01) break;
        }
        if (Math.abs(sum - h.total) >= 0.01) {
          warnings.push(`${name}: "${h.label}" header says ${h.total} but items add to ${sum.toFixed(2)}`);
        }
        items.forEach(i => out.push({ month, category, subcategory, item: i.item, amount: i.amount }));
      });
    });

    // Cross-check against the sheet's own "Total" cell (label in I, amount in J)
    const totalRow = grid.find(row => cleanLabel(row[8]) === "Total" && isNum(row[9]));
    if (totalRow) {
      const mine = out.filter(r => r.month === month).reduce((s, r) => s + r.amount, 0);
      if (Math.abs(mine - totalRow[9]) >= 0.01) {
        warnings.push(`${name}: parsed total ${mine.toFixed(2)} but sheet says ${totalRow[9]}. New category?`);
      }
    }
  });

  return { rows: out, warnings };
}