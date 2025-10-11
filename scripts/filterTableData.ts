function main(
  workbook: ExcelScript.Workbook,
  dataJson: string = "[]",
  rush: string = "rush",            // "rush" | "standard" | "any"
  reqShipDate: string = ""          // "yyyy-MM-dd" from Flow; if empty, fall back to script's "today"
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
    // reqShipDate format: "yyyy-MM-dd"
    const parts = reqShipDate.split("-").map((x) => Number(x));
    const dt = new Date(parts[0], parts[1] - 1, parts[2]);
    targetSerial = Math.floor((dt.getTime() - excelEpoch.getTime()) / msPerDay);
  } else {
    // fallback to the script's "today" (may differ by timezone!)
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    targetSerial = Math.floor((today.getTime() - excelEpoch.getTime()) / msPerDay);
  }

  const wantAny = norm(rush) === "any";

  // filter
  const filtered = (rows as Array<Record<string, unknown>>).filter((r) => {
    const isPriorityOk = wantAny || norm(r["IS Rush?"]) === norm(rush);
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
