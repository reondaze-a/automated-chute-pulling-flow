function main(
    workbook: ExcelScript.Workbook,
    chutesJson: string = "[]"
) {
    // --- helpers ---
    const normChute = (v: unknown) =>
        (v ?? "").toString().replace(/\u00A0/g, " ").trim().substring(0, 4).toUpperCase();

    const parseChute = (ch: string) => {
        const m = /^([A-Za-z])\s*0*?(\d{1,3})$/.exec(ch);
        if (!m) return { zone: "", num: NaN };
        return { zone: m[1].toUpperCase(), num: Number(m[2]) };
    };

    // Line mapping from your sheet
    const rules = [
        // A zone
        { zone: "A", start: 1, end: 152, line: "Line 1" },
        { zone: "A", start: 153, end: 300, line: "Line 2" },
        { zone: "A", start: 301, end: 448, line: "Line 3" },
        { zone: "A", start: 449, end: 600, line: "Line 4" },
        // B zone
        { zone: "B", start: 1, end: 148, line: "Line 7" },
        { zone: "B", start: 149, end: 306, line: "Line 8" },
        { zone: "B", start: 307, end: 450, line: "Line 6" },
        { zone: "B", start: 451, end: 600, line: "Line 5" }
    ];

    const lineOf = (zone: string, num: number) => {
        for (const r of rules) {
            if (r.zone === zone && Number.isFinite(num) && num >= r.start && num <= r.end) {
                return r.line;
            }
        }
        return "";
    };

    // --- parse input ---
    let items: unknown[] = [];
    try { items = JSON.parse(chutesJson) || []; } catch { items = []; }

    // Expect input: [{ CHUTE: "A001" }, ...]
    const source: string[] = (items as Array<Record<string, unknown>>)
        .map((o) => normChute(o["CHUTE"]))
        .filter((v) => v.length > 0);

    // unique + natural sort
    const distinct = Array.from(new Set(source))
        .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

    // bucket into lines
    const buckets: Record<string, string[]> = {};
    const lineOrder = ["Line 1", "Line 2", "Line 3", "Line 4", "Line 5", "Line 6", "Line 7", "Line 8"];
    for (const ln of lineOrder) buckets[ln] = [];

    for (const ch of distinct) {
        const p = parseChute(ch);
        const ln = lineOf(p.zone, p.num);
        if (ln) buckets[ln].push(ch);
    }

    // build matrix: one column per line, rows stack CHUTEs
    const maxLen = lineOrder.reduce((mx, ln) => Math.max(mx, buckets[ln].length), 0);
    const rows: Array<Record<string, string>> = [];
    for (let i = 0; i < maxLen; i++) {
        const row: Record<string, string> = {};
        for (const ln of lineOrder) {
            row[ln] = buckets[ln][i] ?? "";
        }
        rows.push(row);
    }

    return JSON.stringify(rows);
}
