function main(
  workbook: ExcelScript.Workbook,
  dataJson: string = "[]",
  rush: string = "rush",            // used for columnEquals mode or when legacy column exists
  rushColumn: string = "IS Rush?",  // column to inspect (e.g., "Source Code" for CBI)
  reqShipDate: string = "",         // "yyyy-MM-dd" from Flow; if empty, uses script's today
  rushDetect: string = "columnEquals", // "columnEquals" | "columnContains"
  rushSubstring: string = "-EX",        // used only when rushDetect = "columnContains"
  containsMode: string = "include"
) {
  // --- helpers ---
  const msPerDay = 86400000;
  const excelEpoch = new Date(1899, 11, 30); // 1899-12-30

  const norm = (v: unknown) =>
    (v ?? "").toString().replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim().toLowerCase();

  const normalizeChute = (v: unknown) => {
    const t: string = (v ?? "").toString().replace(/\u00A0/g, " ").trim();
    const noTailY: string = t.replace(/[Yy]\s*$/, ""); // drop trailing Y + spaces
    return noTailY.substring(0, 4);
  };

  // parse input rows
  let rows: unknown[] = [];
  try { rows = JSON.parse(dataJson) || []; } catch { rows = []; }

  // target Excel serial for the requested ship date
  let targetSerial: number;
  if (reqShipDate) {
    const parts = reqShipDate.split("-").map((x) => Number(x));
    const dt = new Date(parts[0], parts[1] - 1, parts[2]);
    targetSerial = Math.floor((dt.getTime() - excelEpoch.getTime()) / msPerDay);
  } else {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    targetSerial = Math.floor((today.getTime() - excelEpoch.getTime()) / msPerDay);
  }

  const wantAny = norm(rush) === "any";
  const detect: string = norm(rushDetect);      // "columnequals" | "columncontains"
  const needle: string = norm(rushSubstring);   // e.g., "-ex"
  const mode: string = norm(containsMode); // include | exclude

  // filter
  const filtered = (rows as Array<Record<string, unknown>>).filter((r) => {
    // Priority check
    let isPriorityOk = false;
    if (wantAny) {
      isPriorityOk = true;
    } else {
      const colVal: string = norm(r[rushColumn]);

      if (detect === "columncontains") {
        const has = needle.length === 0 ? true : colVal.includes(needle);
        isPriorityOk = (mode === "exclude") ? !has : has;   // ← invert when exclude
      } else {
        isPriorityOk = colVal === norm(rush);
      }
    }

    const cartonOk = norm(r["Carton Status"]) === "00 - printed";

    const reqShipRaw = Number(r["Req Ship Date"]);
    const reqShip = Number.isFinite(reqShipRaw) ? Math.floor(reqShipRaw) : NaN;
    const reqShipIsTarget = reqShip === targetSerial;

    const chute4 = normalizeChute(r["CHUTE"]);
    const hasChute = chute4.length > 0;

    return isPriorityOk && cartonOk && reqShipIsTarget && hasChute;
  });

  // CHUTE only, dedupe, natural sort
  const unique = Array.from(new Set(filtered.map((r) => normalizeChute(r["CHUTE"]))))
    .filter((v) => v.length > 0)
    .map((v) => ({ CHUTE: v }))
    .sort((a, b) => a.CHUTE.localeCompare(b.CHUTE, undefined, { numeric: true, sensitivity: "base" }));

  return JSON.stringify(unique);
}
