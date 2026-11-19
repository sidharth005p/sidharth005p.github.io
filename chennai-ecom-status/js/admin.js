const CORE = getCore();

const $ = i => document.getElementById(i);

function book(sheets) {
    const wb = XLSX.utils.book_new();
    sheets.forEach(([n, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows :
        [{
            Info: "No data"
        }]), n.slice(0, 31)));
    return wb
}

function fmtTime(v) {
    if (typeof v === "number") {
        const d = new Date(Math.round((v - 25569) * 864e5)),
            p = n => String(n).padStart(2, "0");
        return d.getUTCFullYear() + "-" + p(d.getUTCMonth() + 1) + "-" + p(d.getUTCDate()) + " " + p(d
            .getUTCHours()) + ":" + p(d.getUTCMinutes())
    }
    return String(v || "").slice(0, 16)
}
/* Every Chennai / ROTN ticket in this file counts as a review. */
function parse(buf) {
    const wb = XLSX.read(buf, {
        type: "array"
    });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
        raw: true,
        defval: ""
    });
    if (!rows.length) throw new Error("The file has no rows.");
    const BR = ["chennai", "rotn"];
    const key = n => Object.keys(rows[0]).find(k => k.trim().toLowerCase() === n);
    const kb = key("branch"),
        kt = key("ticket id"),
        kd = key("added time");
    if (!kb || !kt) throw new Error("Columns Branch and Ticket ID are required.");
    const list = rows.filter(r => BR.indexOf(String(r[kb]).trim().toLowerCase()) >= 0).map(r => ({
        t: String(r[kt]).trim(),
        b: String(r[kb]).trim(),
        d: kd ? fmtTime(r[kd]) : ""
    })).filter(r => r.t);
    if (!list.length) throw new Error("No Chennai or ROTN rows found in the file.");
    list.sort((a, b) => a.d.localeCompare(b.d));
    const t = {};
    list.forEach(r => {
        if (t[r.t]) return;
        t[r.t] = {
            d: r.d.slice(0, 10),
            b: r.b
        }
    });
    return {
        asof: list[list.length - 1].d,
        from: list[0].d.slice(0, 10),
        t,
        list
    }
}

function save(name, blob) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove()
}
let nd = null;
$("file").onchange = async e => {
    const fl = e.target.files[0];
    nd = null;
    $("btns").hidden = true;
    $("msg").textContent = "";
    if (!fl) return;
    try {
        nd = parse(await fl.arrayBuffer());
        const n = Object.keys(nd.t).length;
        $("msg").className = "msg g";
        $("msg").textContent = "Read " + n + " Chennai + ROTN tickets from " + nd.from + " to " + nd
            .asof +
            ". Every ticket is compared with all the months of the calls file below; matches count as reviews. ";
        $("btns").hidden = false
    } catch (err) {
        $("msg").className = "msg e";
        $("msg").textContent = err.message || "Could not read the file."
    }
};
$("dj").onclick = () => save("data.json", new Blob([JSON.stringify({
    asof: nd.asof,
    from: nd.from,
    t: nd.t
}, null, 4)], {
    type: "application/json"
}));

let calls = null,
    rep = null;
const MN = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const mname = m => MN[+m.slice(5) - 1] + " " + m.slice(0, 4);
async function makeReport() {
    const m = $("cmsg");
    if (!calls) return;
    $("cbtns").hidden = true;
    try {
        let t, asof, from = "";
        if (nd) {
            t = nd.t;
            asof = nd.asof;
            from = nd.from
        } else {
            const d = await (await fetch("data.json?v=" + Date.now())).json();
            t = d.t;
            asof = d.asof;
            from = d.from || ""
        }
        const month = calls.dateOk ? $("mon").value : "";
        rep = CORE.buildReport(calls.rows, t, {
            range: calls.range,
            asof,
            from,
            month,
            softClosed: $("soft").checked
        });
        const nrev = rep.totals.reviewTickets,
            T = rep.totals;
        m.className = "msg g";
        m.textContent = "Report ready" + (month ? " for " + mname(month) : "") + ": " + rep.franchises
            .length + " franchises. Uploaded review tickets: " + nrev + "; accepted (counted): " + T
            .reviews + "; rejected (not counted): " + rep.rejected.length + ". Denominator: " + T.closed +
            " closed Flipkart/Cloudtail installation tickets" + (month ? " of " + mname(month) : "") +
            " → review % " + T.pct + "%. Estimated technician earnings ₹" + T.earn + "; if all " + nrev +
            " uploaded tickets were accepted: ₹" + T.aEarn + " (+₹" + T.gain + "). " + (calls.dateOk ? "" :
                "WARNING: no call-date column was found in the calls file, so every row of the file is counted as one month. "
                ) + (nd ? "Using the review file chosen above." :
                "Using the review data currently on the website (till " + asof + ").");
        $("cbtns").hidden = false
    } catch (err) {
        m.className = "msg e";
        m.textContent = err.message || "Could not build the report."
    }
}
$("cfile").onchange = async e => {
    const fls = [...e.target.files],
        m = $("cmsg");
    calls = null;
    $("cbtns").hidden = true;
    $("monw").hidden = true;
    m.textContent = "";
    if (!fls.length) return;
    const big = fls.find(f => f.size > 45 * 1024 * 1024);
    if (big) {
        m.className = "msg e";
        m.textContent = "The file " + big.name + " is " + Math.round(big.size / 1048576) +
            " MB, which is too large for a browser to read (it looks like several months of unfiltered calls). Use the reduced file (only e-com installation calls, closed) or download a smaller range.";
        return
    }
    m.className = "msg";
    m.textContent = "Reading the file… please wait.";
    try {
        const parts = [];
        for (const fl of fls) {
            await new Promise(r => setTimeout(r, 50));
            const buf = await fl.arrayBuffer();
            let aoa = [];
            try {
                const wb = XLSX.read(buf, {
                    type: "array",
                    dense: true,
                    sheets: 0
                });
                aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
                    header: 1,
                    raw: true,
                    defval: ""
                })
            } catch (x) {
                aoa = []
            }
            if (!aoa.length) {
                const wb = XLSX.read(buf, {
                    type: "array",
                    sheets: 0
                });
                aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
                    header: 1,
                    raw: true,
                    defval: ""
                })
            }
            parts.push(CORE.pickCalls(aoa))
        }
        calls = {
            rows: [].concat(...parts.map(x => x.rows)),
            range: parts[0].range,
            months: {},
            dateOk: parts.every(x => x.dateOk),
            dateCol: parts[0].dateCol
        };
        parts.forEach(x => Object.keys(x.months).forEach(k => {
            calls.months[k] = (calls.months[k] || 0) + x.months[k]
        }));
        if (calls.dateOk) {
            const ms = Object.keys(calls.months).sort(),
                cur = nd ? nd.asof.slice(0, 7) : ms[ms.length - 1];
            $("mon").innerHTML = ms.map(k =>
                `<option value="${k}">${mname(k)} (${calls.months[k]} calls)</option>`).join("");
            $("mon").value = ms.indexOf(cur) >= 0 ? cur : ms[ms.length - 1];
            $("monw").hidden = false
        }
        await makeReport()
    } catch (err) {
        m.className = "msg e";
        m.textContent = err.message || "Could not read the file."
    }
};
$("mon").onchange = makeReport;
$("dun").onclick = () => {
    const X = x => ({
            "Ticket ID": x.t,
            "Branch": x.b,
            "Franchise code": x.fc,
            "Franchise name": x.fn || "(not in calls file)",
            "Technician": x.tn,
            "Review added": x.d,
            "Reason not counted": x.why,
            "Call status": x.st,
            "Dealer name": x.dl,
            "Origin": x.og,
            "Product": x.pr,
            "Call booked": x.cb,
            "Call type": x.ct,
            "Note": x.note
        }),
        cnt = {};
    rep.rejected.forEach(u => {
        cnt[u.why] = (cnt[u.why] || 0) + 1
    });
    XLSX.writeFile(book([
        ["Summary", Object.keys(cnt).map(k => ({
            Reason: k,
            Tickets: cnt[k]
        }))],
        ["All tickets", rep.rejected.map(X)]
    ]), "Rejected_Review_Tickets.xlsx")
};
$("soft").onchange = makeReport;
$("file").addEventListener("change", () => setTimeout(makeReport, 300));
$("dr").onclick = () => save("report.json", new Blob([JSON.stringify(rep, null, 4)], {
    type: "application/json"
}));
$("dxs").onclick = () => {
    const sum = [],
        flat = [];
    rep.franchises.forEach(f => {
        sum.push({
            "Franchise": f.name,
            "Code": f.code,
            "E-com closed": f.closed,
            "Reviews": f.reviews,
            "Review %": f.pct,
            "Estimated earnings ₹": f.earn,
            "Reviews if all uploaded accepted": f.aReviews,
            "Earnings if all accepted ₹": f.aEarn,
            "Extra if all accepted ₹": f.gain
        });
        CORE.techRows(f, "all", true).forEach(r => flat.push(r))
    });
    XLSX.writeFile(book([
        ["Franchise summary", sum],
        ["Technician slabs", flat]
    ]), "Technician_Slab_Report.xlsx")
};