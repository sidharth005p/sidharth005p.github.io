const CORE = getCore();

const $ = i => document.getElementById(i);
const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;"
} [c]));
const inr = n => "₹" + Number(n || 0).toLocaleString("en-IN"),
  SL = s => s ? s + "%" : "–";
const monthLabel = month => {
  const [year, value] = String(month || "").split("-");
  const names = ["January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December"
  ];
  return names[+value - 1] ? names[+value - 1] + " " + year : month
};
const monthWindow = month => {
  const match = String(month || "").match(/^(\d{4})-(\d{2})$/);
  if (!match) return [];
  const previous = new Date(Date.UTC(+match[1], +match[2] - 2, 1));
  return [previous.getUTCFullYear() + "-" + String(previous.getUTCMonth() + 1).padStart(2, "0"), month]
};
const LB = {
  C1: CORE.CATS.C1.short,
  C2: CORE.CATS.C2.short
};
const reduce = matchMedia("(prefers-reduced-motion:reduce)").matches;
let R = null;

function book(sheets) {
  const wb = XLSX.utils.book_new(),
    used = {};
  sheets.forEach(([n, rows]) => {
    let nm = n.replace(/[\[\]:*?\/\\]/g, " ").slice(0, 31),
      k = 1;
    while (used[nm.toLowerCase()]) {
      nm = n.replace(/[\[\]:*?\/\\]/g, " ").slice(0, 28) + " " + (++k)
    }
    used[nm.toLowerCase()] = 1;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{
      Info: "No data"
    }]), nm)
  });
  return wb
}

function grow(root) {
  requestAnimationFrame(() => requestAnimationFrame(() => root.querySelectorAll("[data-w]").forEach(e => {
    e.style.width = Math.min(100, e.dataset.w) + "%"
  })))
}

function kpis(el, items) {
  el.innerHTML = items.map(([v, l, f]) =>
    `<div class="k"><b data-v="${v}" data-f="${f||""}">0</b><span>${l}</span></div>`).join("");
  el.querySelectorAll("b").forEach(b => {
    const v = +b.dataset.v,
      f = b.dataset.f,
      fmt = x => f === "inr" ? inr(Math.round(x)) : f === "pct" ? (Math.round(x * 100) / 100) + "%" : Math
      .round(x).toLocaleString("en-IN");
    if (reduce) {
      b.textContent = fmt(v);
      return
    }
    const t0 = performance.now();
    (function s(t) {
      const p = Math.min(1, (t - t0) / 700);
      b.textContent = fmt(v * (1 - Math.pow(1 - p, 3)));
      if (p < 1) requestAnimationFrame(s)
    })(t0)
  })
}
const has = (o, q) => !q || Object.values(o).join(" ").toLowerCase().includes(q);
const mk = root => k => root.querySelector('[data-r="' + k + '"]');
const seg = (el, items, cur, fn) => {
  el.innerHTML = items.map(([k, l]) => `<button data-k="${esc(k)}" class="${k===cur?"on":""}">${l}</button>`)
    .join("");
  el.querySelectorAll("button").forEach(b => b.onclick = () => fn(b.dataset.k))
};

/* ============ table with filter arrows (▾) on every column ============ */
const FS = {},
  pop = $("pop");

function closePop() {
  pop.hidden = true;
  pop.innerHTML = ""
}
document.addEventListener("click", e => {
  if (!pop.hidden && !pop.contains(e.target) && !e.target.classList.contains("fa")) closePop()
});
document.addEventListener("keydown", e => {
  if (e.key === "Escape") closePop()
});

function ftable(host, id, cols, rows, redraw, max) {
  max = max || 1500;
  const S = FS[id] || (FS[id] = {
    f: {},
    s: null
  });
  const val = (r, c) => {
      const x = c[4] ? c[4](r) : r[c[1]];
      return String(x == null ? "" : x)
    },
    num = x => parseFloat(String(x).replace(/[^\d.\-]/g, ""));
  let v = rows.filter(r => cols.every(c => !S.f[c[1]] || S.f[c[1]].has(val(r, c))));
  if (S.s) {
    const c = cols.find(c => c[1] === S.s.k);
    if (c) {
      const isn = v.every(r => val(r, c) === "" || !isNaN(num(val(r, c))));
      v = v.slice().sort((a, b) => S.s.d * (isn ? (num(val(a, c)) || 0) - (num(val(b, c)) || 0) : val(a, c)
        .localeCompare(val(b, c))))
    }
  }
  const act = Object.keys(S.f).some(k => S.f[k]);
  host.innerHTML =
    `<p class="note" style="margin:6px 2px">Showing ${Math.min(v.length,max)} of ${rows.length}${act||S.s?` · <button class="lnk" data-clear>Clear filters</button>`:""}</p><div class="tw"><table><thead><tr>${cols.map(c=>`<th class="${c[3]||""}">${c[0]}<button class="fa ${S.f[c[1]]||(S.s&&S.s.k===c[1])?"on":""}" data-k="${esc(c[1])}" aria-label="Filter ${esc(c[0])}">▾</button></th>`).join("")}</tr></thead><tbody>${v.slice(0,max).map(r=>`<tr>${cols.map(c=>`<td class="${c[3]||""}">${c[2]?c[2](r):esc(r[c[1]])}</td>`).join("")}</tr>`).join("")||`<tr><td colspan="${cols.length}" class="empty">No rows match.</td></tr>`}</tbody></table></div>` +
    (v.length > max ? `<p class="note">Showing the first ${max}. Use the Excel download for all rows.</p>` : "");
  const cl = host.querySelector("[data-clear]");
  if (cl) cl.onclick = () => {
    FS[id] = {
      f: {},
      s: null
    };
    redraw()
  };
  host.querySelectorAll(".fa").forEach(b => b.onclick = e => {
    e.stopPropagation();
    const k = b.dataset.k,
      c = cols.find(c => c[1] === k),
      cnt = {};
    rows.forEach(r => {
      const x = val(r, c);
      cnt[x] = (cnt[x] || 0) + 1
    });
    const keys = Object.keys(cnt).sort((a, b) => {
        const na = num(a),
          nb = num(b);
        return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b)
      }).slice(0, 400),
      sel = S.f[k];
    pop.innerHTML =
      `<div class="row"><button data-s="1">Sort A→Z</button><button data-s="-1">Sort Z→A</button></div><input type="search" placeholder="Search values" aria-label="Search values"><div class="lst">${keys.map(x=>`<label><input type="checkbox" value="${esc(x)}" ${!sel||sel.has(x)?"checked":""}> <span>${esc(x)||"(blank)"} <em style="color:var(--mu);font-style:normal">(${cnt[x]})</em></span></label>`).join("")}</div><div class="row"><button data-all>Select all</button><button data-none>Clear</button><button class="pr" data-ok>Apply</button></div>`;
    const rc = b.getBoundingClientRect();
    pop.hidden = false;
    pop.style.left = Math.max(8, Math.min(rc.left, innerWidth - 258)) + "px";
    pop.style.top = Math.min(rc.bottom + 4, innerHeight - 330) + "px";
    const boxes = () => [...pop.querySelectorAll(".lst input")];
    pop.querySelector('input[type=search]').oninput = ev => {
      const q = ev.target.value.toLowerCase();
      pop.querySelectorAll(".lst label").forEach(l => l.hidden = !l.textContent.toLowerCase().includes(q))
    };
    pop.querySelector("[data-all]").onclick = () => boxes().forEach(x => x.checked = true);
    pop.querySelector("[data-none]").onclick = () => boxes().forEach(x => x.checked = false);
    pop.querySelectorAll("[data-s]").forEach(x => x.onclick = () => {
      S.s = {
        k,
        d: +x.dataset.s
      };
      closePop();
      redraw()
    });
    pop.querySelector("[data-ok]").onclick = () => {
      const on = boxes().filter(x => x.checked).map(x => x.value);
      if (on.length === keys.length) delete S.f[k];
      else S.f[k] = new Set(on);
      closePop();
      redraw()
    }
  })
}

/* ============ dashboards (main / new rules / all) ============ */
function mountDash(root, mode, report) {
  const D = report || R,
    all = mode === "all",
    rulesOnly = mode === "new",
    $$ = mk(root),
    vis = x => all || x.closed || x.reviews;
  let fsel = "",
    rcat = "all";
  const meta = R.newRulesMeta || {
    months: (R.nr && R.nr.window && R.nr.window.months) || monthWindow(R.month || String(R.asof || "").slice(0, 7)),
    included: D.totals.reviewTickets,
    excluded: Math.max(0, R.totals.reviewTickets - D.totals.reviewTickets)
  };
  const rejBy = {};
  D.rejected.forEach(x => {
    if (x.fc) rejBy[x.fc] = (rejBy[x.fc] || 0) + 1
  });
  root.innerHTML =
    `<p class="scope" data-r="scope"></p>${rulesOnly?`<p class="pgnote">New acceptance rule: only tickets uploaded in ${meta.months.map(monthLabel).join(" and ")} are eligible. Tickets uploaded outside this period or without an upload date are excluded, not accepted. Included ${meta.included} of ${R.totals.reviewTickets} uploaded tickets; ${meta.excluded} excluded.</p>`:""}<section class="kpis" data-r="kp"></section>${all?"":'<section class="kpis k3" data-r="ks"></section>'}${all?`<p class="pgnote">The case where <b>every uploaded review ticket (${D.totals.reviewTickets}) is accepted</b>: the rejected tickets (cancelled calls, other dealers) are added to the reviews of their franchise and technician; the closed-ticket denominator stays the same. ${D.totals.unattr} ticket${D.totals.unattr===1?"":"s"} with no call record cannot be placed. Extra money = earnings if all accepted − earnings now.</p>`:""}
<section data-r="ov"><div class="tools"><input data-r="q" type="search" placeholder="Search franchise name or code" aria-label="Search franchise"><select data-r="so" aria-label="Sort"><option value="pct">Sort: Review %</option><option value="earn">Sort: Earnings</option>${all?'<option value="gain">Sort: Extra money</option>':""}<option value="closed">Sort: Tickets closed</option><option value="name">Sort: Name</option></select><select data-r="fb" aria-label="Filter by review %"><option value="">Filter: all franchises</option><option value="75">Review % 75 and above</option><option value="50">Review % 50 and above</option><option value="25">Review % 25 and above</option><option value="low">Review % below 25</option><option value="none">No reviews yet</option></select></div><div data-r="fl"></div></section>
<section data-r="dv" hidden><button class="bk" data-r="bk">‹ All franchises</button><div class="tools"><div class="seg" data-r="seg"></div><input data-r="tq" type="search" placeholder="Search technician" aria-label="Search technician"><select data-r="ts" aria-label="Filter by slab"><option value="">Filter: all slabs</option><option value="0">No slab yet</option><option value="25">25% slab</option><option value="50">50% slab</option><option value="75">75% slab</option><option value="100">100% slab</option></select></div><div class="tl" data-r="tl"></div></section>`;
  const FV = f => all ? {
    pct: f.aPct,
    reviews: f.aReviews,
    earn: f.aEarn,
    gain: f.gain
  } : {
    pct: f.pct,
    reviews: f.reviews,
    earn: f.earn,
    gain: 0
  };
  const SORT = {
    pct: (a, b) => FV(b).pct - FV(a).pct || b.closed - a.closed,
    earn: (a, b) => FV(b).earn - FV(a).earn,
    gain: (a, b) => b.gain - a.gain,
    closed: (a, b) => b.closed - a.closed,
    name: (a, b) => a.name.localeCompare(b.name)
  };
  const BAND = {
    "": () => true,
    75: f => FV(f).pct >= 75,
    50: f => FV(f).pct >= 50,
    25: f => FV(f).pct >= 25,
    low: f => FV(f).pct < 25,
    none: f => FV(f).reviews === 0
  };

  function summary(el, up, ac, rj) {
    kpis(el, [
      [up, "Review tickets uploaded"],
      [ac, "Accepted (counted as reviews)"],
      [rj, "Rejected (not counted)"]
    ])
  }

  function show(code) {
    fsel = code;
    rcat = "all";
    $$("tq").value = "";
    $$("ts").value = "";
    $$("ov").hidden = !!code;
    $$("dv").hidden = !code;
    const t = D.totals;
    if (!code) {
      $$("scope").textContent = all ? "If all " + t.reviewTickets +
        " uploaded tickets are accepted – all franchises" : rulesOnly ?
        "Dashboard based on new rules – all franchises" : "All franchises";
      kpis($$("kp"), all ? [
        [t.reviewTickets, "Uploaded tickets accepted"],
        [t.aPct, "Review %", "pct"],
        [t.aEarn, "Earnings if all accepted", "inr"],
        [t.gain, "Extra money vs now", "inr"]
      ] : [
        [t.closed, "E-com installations closed"],
        [t.reviews, "Accepted reviews"],
        [t.pct, "Review %", "pct"],
        [t.earn, "Estimated technician earnings", "inr"]
      ]);
      if (!all) summary($$("ks"), t.reviewTickets, t.reviews, D.rejected.length);
      ov()
    } else {
      const f = D.franchises.find(x => x.code === code);
      $$("scope").textContent = f.name;
      kpis($$("kp"), all ? [
        [f.aReviews, "Reviews if all accepted"],
        [f.aPct, "Review %", "pct"],
        [f.aEarn, "Earnings if all accepted", "inr"],
        [f.gain, "Extra money vs now", "inr"]
      ] : [
        [f.closed, "E-com installations closed"],
        [f.reviews, "Accepted reviews"],
        [f.pct, "Review %", "pct"],
        [f.earn, "Technician earnings", "inr"]
      ]);
      if (!all) summary($$("ks"), f.reviews + (rejBy[f.code] || 0), f.reviews, rejBy[f.code] || 0);
      det();
      scrollTo({
        top: 0,
        behavior: reduce ? "auto" : "smooth"
      })
    }
  }

  function ov() {
    const q = $$("q").value.trim().toLowerCase(),
      b = BAND[$$("fb").value];
    const l = D.franchises.filter(f => vis(f) && b(f) && (!q || (f.name + " " + f.code).toLowerCase().includes(q)))
      .sort(SORT[$$("so").value]);
    $$("fl").innerHTML = l.length ? l.map((f, i) => {
      const v = FV(f);
      return `<button class="fr" data-c="${esc(f.code)}"><span class="rk">${i+1}</span><span class="nm"><b>${esc(f.name)}</b><small>Code ${esc(f.code)} · ${f.techs.length} technician${f.techs.length>1?"s":""}</small></span><span class="bar"><i data-w="${v.pct}"></i></span><span class="pc">${v.pct}%</span><span class="nu">${v.reviews}/${f.closed}</span><span class="er">${inr(v.earn)}${all&&f.gain?`<em class="gn">+${inr(f.gain)}</em>`:""}</span></button>`
    }).join("") : '<div class="empty">No franchise matches your search or filter.</div>';
    $$("fl").querySelectorAll(".fr").forEach(b => b.onclick = () => show(b.dataset.c));
    grow(root)
  }

  function ladder(x, label) {
    const r = all ? x.a : x,
      n = r.next,
      s = r.slab;
    const nxt = n ?
      `Next slab <b>${n}%</b>: <b>${r.need} more review${r.need>1?"s":""}</b> → earnings ${inr(r.nextEarn)} <b style="color:var(--teal)">(+${inr(r.nextEarn-r.earn)})</b>` :
      "Top slab reached.";
    const head = all ? (x.extra ?
      `<b class="pos">+${x.extra} rejected ticket${x.extra>1?"s":""} accepted.</b> Slab ${SL(x.slab)||"none"} → ${SL(s)||"none"}, earnings ${inr(x.earn)} → <b>${inr(r.earn)}</b> (<span class="pos">+${inr(r.earn-x.earn)}</span>). ` :
      "No rejected ticket in this category. ") : "";
    return `<div class="cg"><div class="ch"><span>${esc(label)}</span><em class="chip ${s?(n?"s":"t"):""}">${s?s+"% slab · ₹"+r.rate+" per review":"Below 25% slab"}</em></div>
<div class="trk"><i data-w="${r.pct}"></i>${[25,50,75].map(p=>`<u style="left:${p}%"></u>`).join("")}</div>
<div class="lb"><span style="left:0;transform:none">0</span>${[25,50,75,100].map(p=>`<span style="left:${p}%">${p}%</span>`).join("")}</div>
<p class="nx">${r.reviews} reviews vs ${x.closed} closed (${r.pct}%) · earned <b>${inr(r.earn)}</b>. ${head}${nxt}</p></div>`
  }

  function det() {
    const f = D.franchises.find(x => x.code === fsel),
      q = $$("tq").value.trim().toLowerCase(),
      sf = $$("ts").value;
    seg($$("seg"), [
      ["all", "Both categories"],
      ["C1", LB.C1],
      ["C2", LB.C2]
    ], rcat, k => {
      rcat = k;
      det()
    });
    const csOf = t => CORE.ORDER.filter(c => t.cats[c] && vis(t.cats[c]) && (rcat === "all" || rcat === c));
    const l = f.techs.filter(t => vis(t) && csOf(t).length && (!q || t.name.toLowerCase().includes(q)) && (sf ===
      "" || csOf(t).some(c => String((all ? t.cats[c].a : t.cats[c]).slab) === sf)));
    $$("tl").innerHTML = l.length ? l.map(t => {
      const cs = csOf(t);
      const sc = cs.reduce((a, c) => {
        const x = t.cats[c],
          r = all ? x.a : x;
        return {
          closed: a.closed + x.closed,
          reviews: a.reviews + r.reviews,
          earn: a.earn + r.earn,
          gain: a.gain + (x.a.earn - x.earn)
        }
      }, {
        closed: 0,
        reviews: 0,
        earn: 0,
        gain: 0
      });
      return `<article class="tc"><div class="th"><div><h3>${esc(t.name)}</h3><small>${sc.reviews} reviews · ${sc.closed} closed tickets</small></div><div class="te"><span>${all?"Earnings if all accepted":"Earned"}</span><b>${inr(sc.earn)}</b>${all&&sc.gain?`<span class="pos">+${inr(sc.gain)} extra</span>`:""}</div></div>${cs.map(c=>ladder(t.cats[c],CORE.CATS[c].label)).join("")}</article>`
    }).join("") : '<div class="empty">No technician matches your search or filter.</div>';
    grow(root)
  }
  $$("q").oninput = ov;
  $$("so").onchange = ov;
  $$("fb").onchange = ov;
  $$("tq").oninput = () => fsel && det();
  $$("ts").onchange = () => fsel && det();
  $$("bk").onclick = () => show("");
  return {
    open() {
      show(fsel)
    }
  }
}

/* ============ rejected tickets ============ */
const SHORT = w => w === CORE.REASONS.NF ? "Not in calls file" : w === CORE.REASONS.DEALER ? "Other dealer" : w ===
  CORE.REASONS.CANC ? "Cancelled" : w.slice(0, 28);
const RCOLS = [
  ["Ticket ID", "t"],
  ["Branch", "b"],
  ["Franchise code", "fc"],
  ["Franchise name", "fn"],
  ["Technician", "tn"],
  ["Review added", "d"],
  ["Reason not counted", "why"],
  ["Call status", "st"],
  ["Dealer name", "dl"],
  ["Origin", "og"],
  ["Product", "pr"],
  ["Call booked", "cb"],
  ["Call type", "ct"],
  ["Note", "note"]
];
const XROW = x => ({
  "Ticket ID": x.t,
  "Branch": x.b,
  "Franchise code": x.fc || "",
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
});

function byFranchise(rows) {
  const m = {};
  rows.forEach(x => {
    const k = (x.fc || "") + "|" + (x.fn || "(not in calls file)") + "|" + x.b;
    m[k] = (m[k] || 0) + 1
  });
  return Object.keys(m).map(k => {
    const p = k.split("|");
    return {
      code: p[0],
      name: p[1],
      branch: p[2],
      n: m[k]
    }
  }).sort((a, b) => b.n - a.n || a.name.localeCompare(b.name))
}
const REASONS = () => {
  const all = R.rejected;
  return CORE.REASON_ORDER.filter(w => all.some(x => x.why === w)).concat([...new Set(all.map(x => x.why))]
    .filter(w => CORE.REASON_ORDER.indexOf(w) < 0))
};

function mountRej(root) {
  const $$ = mk(root),
    all = R.rejected,
    t = R.totals,
    reasons = REASONS();
  let sub = "sum";
  root.innerHTML =
    `<p class="scope">Rejected tickets</p><section class="kpis" data-r="kp"></section><p class="pgnote r">Review tickets that were uploaded but are <b>not counted</b> in the review dashboard, with the reason, franchise and technician. Use the ▾ arrow on any column to filter or sort, and the search bar to find a ticket.</p>
<div class="tools"><div class="seg" data-r="seg"></div><input data-r="q" type="search" placeholder="Search ticket, franchise, technician, dealer…" aria-label="Search"></div><div data-r="body"></div>`;
  kpis($$("kp"), [
    [t.reviewTickets, "Review tickets uploaded"],
    [t.reviews, "Accepted (counted)"],
    [all.length, "Rejected / not counted"],
    [reasons.length ? all.filter(x => x.why === reasons[0]).length : 0, reasons[0] ? reasons[0].split(" (")[0] :
      "–"
    ]
  ]);
  const items = [
    ["sum", "Summary"],
    ["fr", "By franchise"],
    ["all", "All (" + all.length + ")"]
  ].concat(reasons.map((w, i) => ["r" + i, SHORT(w) + " (" + all.filter(x => x.why === w).length + ")"]));

  function draw() {
    const q = $$("q").value.trim().toLowerCase();
    seg($$("seg"), items, sub, k => {
      sub = k;
      draw()
    });
    const b = $$("body");
    if (sub === "sum") {
      const rc = reasons.map(w => ({
          w,
          n: all.filter(x => x.why === w).length
        })).filter(x => has(x, q)),
        bc = {};
      all.forEach(x => {
        bc[x.b] = (bc[x.b] || 0) + 1
      });
      const bl = Object.keys(bc).map(k => ({
        b: k,
        n: bc[k]
      })).filter(x => has(x, q));
      b.innerHTML =
        `<div class="two"><div><p class="sub2">By reason</p><div data-r="t1"></div></div><div><p class="sub2">By branch</p><div data-r="t2"></div></div></div><p class="note">Total rejected / not counted: <b>${all.length}</b> of ${t.reviewTickets} uploaded.</p>`;
      ftable($$("t1"), "rej-sum-r", [
        ["Reason", "w"],
        ["Tickets", "n", null, "n"]
      ], rc, draw);
      ftable($$("t2"), "rej-sum-b", [
        ["Branch", "b"],
        ["Tickets", "n", null, "n"]
      ], bl, draw)
    } else if (sub === "fr") {
      b.innerHTML = '<div data-r="t1"></div>';
      ftable($$("t1"), "rej-fr", [
        ["Franchise code", "code"],
        ["Franchise name", "name"],
        ["Branch", "branch"],
        ["Unmatched tickets", "n", null, "n"]
      ], byFranchise(all).filter(x => has(x, q)), draw)
    } else {
      const w = sub === "all" ? null : reasons[+sub.slice(1)];
      b.innerHTML = '<div data-r="t1"></div>';
      ftable($$("t1"), "rej-" + sub, RCOLS, all.filter(x => (!w || x.why === w) && has(x, q)), draw)
    }
  }
  $$("q").oninput = draw;
  return {
    open: draw
  }
}

function mountInsights(root) {
  root.innerHTML = `<p class="scope">Rejection insights</p><p class="note">Loading rejection analysis…</p>`;

  function draw(data) {
    const sheets = data.sheets || {};
    const reasons = (sheets["Top 10 - Chennai + ROTN"] || []).slice(4).filter(r => r && r.length >= 7);
    const reasonRows = reasons.filter(r => /^\d+$/.test(String(r[0]))).map(r => ({
      rank: r[0],
      name: r[1],
      count: r[3],
      pct: r[4] * 100,
      action: r[6]
    }));
    const monthly = ((sheets["Rejection rate by month"] || []).slice(4).filter(r => r && r.length >= 8));
    const platform = ((sheets["By platform"] || []).slice(-2).filter(r => r && r[0] === "Amazon (Cloudtail)" || r && r[0] === "Flipkart"));
    const franchises = (sheets["Chennai-ROTN franchises"] || []).slice(4).filter(r => r && r.length >= 7).slice(0, 10);
    const topActions = reasonRows.slice(0, 5);
    const barWidth = pct => Math.max(4, Math.min(100, pct));
    root.innerHTML = `
      <p class="scope">Rejection insights</p>
      <p class="note">Rating rejection analysis · Chennai &amp; ROTN · May–October 2026</p>
      <section class="kpis insight-kpis">
        <div class="k"><b>5,170</b><span>reviews submitted</span></div>
        <div class="k"><b>1,924</b><span>rejected</span></div>
        <div class="k"><b>37.2%</b><span>combined rejected %</span></div>
        <div class="k"><b>1,659</b><span>tickets with a recorded reason</span></div>
      </section>

      <div class="insight-grid three">
        <article class="insight-card"><span class="insight-stat">631 tickets · 38%</span><h3>Customer rates the wrong thing</h3><p>Delivery, seller or service rating instead of the product rating.</p></article>
        <article class="insight-card"><span class="insight-stat">893 tickets · 53.8%</span><h3>Proof upload mistakes</h3><p>Bill copy or rating screenshot missing or unclear, or the Order ID typed incorrectly.</p></article>
        <article class="insight-card"><span class="insight-stat">133 tickets · 8%</span><h3>Process issues</h3><p>Ticket not registered in CRM, or registered in an earlier month.</p></article>
      </div>

      <section class="insight-section">
        <h3>Top five reasons · 84% of recorded rejections</h3>
        <div class="reason-list">
          ${topActions.map(r => `
            <article class="reason-row">
              <div class="reason-rank">#${r.rank}</div>
              <div class="reason-name"><strong>${esc(r.name)}</strong><span>${r.count.toLocaleString("en-IN")} tickets</span></div>
              <div class="reason-pct">${r.pct.toFixed(1)}%</div>
              <div class="reason-bar"><i style="width:${barWidth(r.pct)}%"></i></div>
            </article>`).join("")}
        </div>
      </section>

      <section class="insight-section">
        <h3>How to stop the top five reasons</h3>
        <div class="tw"><table class="action-table"><thead><tr><th>#</th><th>Reason</th><th>Tickets</th><th>What to do</th></tr></thead><tbody>
          ${topActions.map(r => `<tr><td>${r.rank}</td><td>${esc(r.name)}</td><td>${r.count.toLocaleString("en-IN")} (${r.pct.toFixed(1)}%)</td><td>${esc(r.action)}</td></tr>`).join("")}
        </tbody></table></div>
      </section>

      <section class="insight-section">
        <h3>Rejected share by month</h3>
        <div class="trend-table">
          ${monthly.map(r => `
            <article class="trend-row">
              <strong>${esc(r[0])}</strong>
              <div class="trend-cell"><span>${(r[3] * 100).toFixed(1)}% All</span><div class="trend-bar"><i style="width:${barWidth(r[3] * 100)}%"></i></div></div>
              <div class="trend-cell"><span>${(r[6] * 100).toFixed(1)}% Chennai</span><div class="trend-bar"><i style="width:${barWidth(r[6] * 100)}%"></i></div></div>
              <div class="trend-cell"><span>${(r[9] * 100).toFixed(1)}% ROTN</span><div class="trend-bar"><i style="width:${barWidth(r[9] * 100)}%"></i></div></div>
            </article>`).join("")}
        </div>
        <p class="insight-note"><b>Trend:</b> Chennai peaked at 45.3% in October after rising for three months. ROTN eased to 39.0% in October.</p>
      </section>

      <section class="insight-section">
        <h3>Platform impact</h3>
        <div class="platform-compare">
          ${platform.map(r => `<article class="insight-card"><span class="insight-stat">${esc(r[0])}</span><b>${(r[3] * 100).toFixed(1)}%</b><p>${r[2].toLocaleString("en-IN")} rejected of ${r[1].toLocaleString("en-IN")} submitted</p></article>`).join("")}
        </div>
      </section>

      <section class="insight-section">
        <h3>Franchises to coach first</h3>
        <div class="coach-list">
          <div class="coach-head"><span>Franchise</span><span>Submitted</span><span>Rejected %</span><span>Main rejection reason</span></div>
          ${franchises.filter(r => r[3] >= 20 && !/^manivannan\s*s/i.test(String(r[1])) && !/^raja\s*m/i.test(String(r[1]))).slice(0, 5).map(r => `
            <article class="coach-row">
              <strong>${esc(r[1])} · ${esc(r[0])}</strong>
              <span><small>Submitted</small><b>${r[2].toLocaleString("en-IN")}</b></span>
              <span><small>Rejected</small><b class="trend-pct">${(r[4] * 100).toFixed(1)}%</b></span>
              <span class="coach-reason"><small>Main rejection reason</small>${esc(r[5])}<br><b>${r[6].toLocaleString("en-IN")} tickets</b></span>
            </article>`).join("")}
        </div>
        <p class="note">Franchises with at least 20 submissions, June–October 2026.</p>
      </section>

      <div class="insight-note"><b>CRM check:</b> 110 Chennai + ROTN tickets were not registered in CRM. This is an e-commerce team process issue, not a technician fault. Check e-commerce orders in CRM before the visit and escalate missing tickets the same day.</div>
      <p class="note">Source: Rating_Rejection_Analysis.xlsx · reason counts cover June–October 2026; May has no reason column.</p>`;
  }

  fetch("analysis.json?v=" + Date.now()).then(r => {
    if (!r.ok) throw new Error("Analysis not available");
    return r.json();
  }).then(draw).catch(() => {
    root.innerHTML = `<p class="scope">Rejection insights</p><div class="insight-note">The rejection analysis file is not available yet.</div>`;
  });
  return {
    open() {}
  }
}

function mountProjectedLoss(root) {
  root.innerHTML = `<p class="scope">Projected loss based on rejection rate</p><p class="note">Loading projection…</p>`;

  fetch("analysis.json?v=" + Date.now()).then(r => {
    if (!r.ok) throw new Error("Analysis not available");
    return r.json();
  }).then(data => {
    const sheets = data.sheets || {};
    const normalizeName = value => String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const franchiseHistory = new Map();
    (sheets["Chennai-ROTN franchises"] || []).slice(4).forEach(r => {
      if (r && r[1]) franchiseHistory.set(normalizeName(r[1]), r);
    });
    const monthly = (sheets["Rejection rate by month"] || []).slice(4).filter(r => r && r.length >= 10);
    const october = monthly.find(r => ["oct", "october"].includes(String(r[0]).toLowerCase()));
    const octoberRates = {
      Chennai: october ? october[6] * 100 : 45.2768729,
      ROTN: october ? october[9] * 100 : 38.9527458,
      All: october ? october[3] * 100 : 43.3185967
    };
    const nearest100 = value => Math.round((value || 0) / 100) * 100;
    const currentFranchises = (R && R.franchises ? R.franchises : []).slice();

    function categorySlabLoss(cats, projectedRejected) {
      const names = ["C1", "C2"];
      const totalClosed = names.reduce((sum, c) => sum + (cats[c] ? cats[c].closed || 0 : 0), 0);
      let allocated = 0;
      let loss = 0;
      let remainingEarn = 0;
      names.forEach((c, i) => {
        const cat = cats[c];
        if (!cat) return;
        const categoryClosed = cat.closed || 0;
        const projected = i === 0
          ? (totalClosed ? Math.round(projectedRejected * categoryClosed / totalClosed) : 0)
          : Math.max(0, projectedRejected - allocated);
        allocated += projected;
        const currentReviews = cat.reviews || 0;
        const currentEarn = cat.earn || 0;
        const remainingReviews = Math.max(0, currentReviews - projected);
        const remainingPct = categoryClosed ? remainingReviews / categoryClosed * 100 : (remainingReviews ? 100 : 0);
        const slab = CORE.slabOf(remainingPct);
        const rate = CORE.rateOf(c, slab);
        const categoryRemainingEarn = rate * remainingReviews;
        remainingEarn += categoryRemainingEarn;
        loss += Math.max(0, currentEarn - categoryRemainingEarn);
      });
      return {
        loss,
        remainingEarn
      };
    }

    function aggregateFranchiseCategories(f) {
      const cats = {
        C1: { closed: 0, reviews: 0, earn: 0 },
        C2: { closed: 0, reviews: 0, earn: 0 }
      };
      (f.techs || []).forEach(t => {
        ["C1", "C2"].forEach(c => {
          if (!t.cats || !t.cats[c]) return;
          cats[c].closed += t.cats[c].closed || 0;
          cats[c].reviews += t.cats[c].reviews || 0;
          cats[c].earn += t.cats[c].earn || 0;
        });
      });
      return cats;
    }

    const franchiseRows = currentFranchises.map(f => {
      const historical = franchiseHistory.get(normalizeName(f.name));
      const branch = historical && String(historical[0]).toLowerCase().startsWith("rotn") ? "ROTN" : "Chennai";
      const rejectionPct = octoberRates[branch];
      const projectedRejected = rejectionPct == null ? 0 : Math.round(f.closed * rejectionPct / 100);
      const earnings = f.earn || 0;
      const earningsAfterLoss = nearest100(earnings * (1 - rejectionPct / 100));
      const slabLoss = nearest100(Math.max(0, earnings - earningsAfterLoss));
      return {
        name: f.name,
        code: f.code,
        branch,
        rejectionPct,
        closed: f.closed || 0,
        projectedRejected,
        earnings,
        earningsAfterLoss,
        loss: slabLoss
      };
    }).sort((a, b) => b.loss - a.loss);

    const technicianRows = currentFranchises.flatMap(f => {
      const historical = franchiseHistory.get(normalizeName(f.name));
      const branch = historical && String(historical[0]).toLowerCase().startsWith("rotn") ? "ROTN" : "Chennai";
      const rejectionPct = octoberRates[branch];
      return (f.techs || []).map(t => {
        const projectedRejected = rejectionPct == null ? 0 : Math.round((t.closed || 0) * rejectionPct / 100);
        const earnings = t.earn || 0;
        const cats = {
          C1: t.cats && t.cats.C1 ? {
            closed: t.cats.C1.closed || 0,
            reviews: t.cats.C1.reviews || 0,
            earn: t.cats.C1.earn || 0
          } : null,
          C2: t.cats && t.cats.C2 ? {
            closed: t.cats.C2.closed || 0,
            reviews: t.cats.C2.reviews || 0,
            earn: t.cats.C2.earn || 0
          } : null
        };
        const earningsAfterLoss = nearest100(earnings * (1 - rejectionPct / 100));
        const slabLoss = nearest100(Math.max(0, earnings - earningsAfterLoss));
        return {
          name: t.name,
          franchise: f.name,
          branch,
          rejectionPct,
          projectedRejected,
          earnings,
          earningsAfterLoss,
          loss: slabLoss
        };
      });
    }).sort((a, b) => b.loss - a.loss);

    const projectedLossTotal = nearest100(franchiseRows.reduce((sum, r) => sum + r.loss, 0));
    const projectedEarningsAfterLoss = nearest100(franchiseRows.reduce((sum, r) => sum + r.earningsAfterLoss, 0));
    const currentTotals = R && R.totals ? R.totals : {};

    root.innerHTML = `
      <p class="scope">Projected loss based on rejection rate</p>
      <p class="note">Projection uses the October 2026 Chennai and ROTN rejection rates only, not the multi-month average. Formula: <b>earnings after loss = earnings now × (1 − rejection %)</b>, rounded to the nearest ₹100; <b>projected loss = earnings now − earnings after loss</b>, also rounded to the nearest ₹100.</p>

      <section class="kpis insight-kpis">
        <div class="k"><b>${(currentTotals.closed || 0).toLocaleString("en-IN")}</b><span>E-com installations closed</span></div>
        <div class="k"><b>${(currentTotals.reviews || 0).toLocaleString("en-IN")}</b><span>Accepted reviews</span></div>
        <div class="k"><b>${currentTotals.pct || 0}%</b><span>Review %</span></div>
        <div class="k"><b>${inr(currentTotals.earn || 0)}</b><span>Estimated technician earnings</span></div>
      </section>
      <section class="kpis insight-kpis">
        <div class="k"><b>${(currentTotals.reviewTickets || 0).toLocaleString("en-IN")}</b><span>Review tickets uploaded</span></div>
        <div class="k"><b>${(currentTotals.reviews || 0).toLocaleString("en-IN")}</b><span>Accepted (counted as reviews)</span></div>
        <div class="k"><b>${(currentTotals.unmatched || 0).toLocaleString("en-IN")}</b><span>Rejected / not counted</span></div>
      </section>
      <section class="kpis insight-kpis">
        <div class="k"><b>${inr(projectedLossTotal)}</b><span>Projected earnings loss</span></div>
        <div class="k"><b>${inr(projectedEarningsAfterLoss)}</b><span>Projected earnings after loss</span></div>
      </section>
      <div class="insight-note"><b>Advisory:</b> This loss calculation is based strictly on the October 2026 rejection rate. The projected loss may be lower if rating-submission rejections are reduced.</div>

      <section class="insight-section">
        <h3>Projected loss by franchise</h3>
        <div id="projected-franchise-table"></div>
      </section>

      <section class="insight-section">
        <h3>Projected loss by technician</h3>
        <div id="projected-technician-table"></div>
      </section>`;

    const draw = () => {
      ftable(root.querySelector("#projected-franchise-table"), "projected-franchise", [
        ["Franchise", "name", r => `<strong>${esc(r.name)}</strong><br><small>${esc(r.code)}</small>`],
        ["October rejection %", "rejectionPct", r => r.rejectionPct == null ? "—" : r.rejectionPct.toFixed(1) + "%"],
        ["Current closed tickets", "closed", r => r.closed.toLocaleString("en-IN")],
        ["Earnings now", "earnings", r => inr(r.earnings)],
        ["Earnings after projected loss", "earningsAfterLoss", r => inr(r.earningsAfterLoss)],
        ["Projected earnings loss", "loss", r => `<span class="loss">${inr(r.loss)}</span>`, "n"]
      ], franchiseRows, draw);

      ftable(root.querySelector("#projected-technician-table"), "projected-technician", [
        ["Technician", "name", r => `<strong>${esc(r.name)}</strong>`],
        ["Franchise", "franchise"],
        ["Earnings now", "earnings", r => inr(r.earnings)],
        ["Earnings after projected loss", "earningsAfterLoss", r => inr(r.earningsAfterLoss)],
        ["Projected earnings loss", "loss", r => `<span class="loss">${inr(r.loss)}</span>`, "n"]
      ], technicianRows, draw);
    };
    draw();
  }).catch(() => {
    root.innerHTML = `<p class="scope">Projected loss based on rejection rate</p><div class="insight-note">The rejection analysis file is not available yet.</div>`;
  });

  return {
    open() {}
  };
}

/* ============ Excel download of selected subsections ============ */
function dlItems() {
  const all = R.rejected,
    reasons = REASONS(),
    rejBy = {};
  all.forEach(x => {
    if (x.fc) rejBy[x.fc] = (rejBy[x.fc] || 0) + 1
  });
  const vis = x => x.closed || x.reviews;
  const I = [
    ["main", "dash_fr", "Franchise summary (review %, earnings)", "Franchise summary", () => R.franchises.filter(
      vis).map(f => ({
      "Franchise": f.name,
      "Code": f.code,
      "Technicians": f.techs.length,
      "E-com closed": f.closed,
      "Accepted reviews": f.reviews,
      "Review %": f.pct,
      "Estimated earnings ₹": f.earn,
      "Rejected tickets (not counted)": rejBy[f.code] || 0
    }))],
    ["main", "dash_tech", "Technician slabs and earnings (all franchises)", "Technician slabs", () => R.franchises
      .flatMap(f => CORE.techRows(f, "all", false))
    ],
    ["rej", "rej_sum", "Rejected – Summary (by reason)", "Rejected - Summary", () => reasons.map(w => ({
      Reason: w,
      Tickets: all.filter(x => x.why === w).length
    })).concat([{
      Reason: "Total",
      Tickets: all.length
    }])],
    ["rej", "rej_fr", "Rejected – By franchise", "Rejected - By franchise", () => byFranchise(all).map(x => ({
      "Franchise code": x.code,
      "Franchise name": x.name,
      "Branch": x.branch,
      "Unmatched tickets": x.n
    }))],
    ["rej", "rej_all", "Rejected – All tickets (" + all.length + ")", "Rejected - All tickets", () => all.map(
      XROW)]
  ];
  reasons.forEach((w, i) => I.push(["rej", "rej_" + i, "Rejected – " + SHORT(w) + " (" + all.filter(x => x.why ===
    w).length + ")", "Rejected - " + SHORT(w), () => all.filter(x => x.why === w).map(XROW)]));
  I.push(["all", "all_fr", "If all accepted – Franchise summary", "All accepted - Franchises", () => R.franchises
    .filter(f => f.closed || f.reviews || f.extra).map(f => ({
      "Franchise": f.name,
      "Code": f.code,
      "E-com closed": f.closed,
      "Reviews now": f.reviews,
      "Review % now": f.pct,
      "Earnings now ₹": f.earn,
      "Reviews if all accepted": f.aReviews,
      "Review % if all accepted": f.aPct,
      "Earnings if all accepted ₹": f.aEarn,
      "Extra ₹": f.gain
    }))
  ]);
  I.push(["all", "all_tech", "If all accepted – Technician details", "All accepted - Technicians", () => R
    .franchises.flatMap(f => CORE.techRows(f, "all", true))
  ]);
  return I.map(([g, id, label, sheet, rows]) => ({
    g,
    id,
    label,
    sheet,
    rows
  }))
}
const GL = {
  main: "Review dashboard",
  rej: "Rejected tickets",
  all: "If all uploaded tickets are accepted"
};

function openDl() {
  const I = dlItems();
  let h = "";
  ["main", "rej", "all"].forEach(g => {
    h +=
      `<fieldset><legend>${GL[g]}</legend>${I.filter(x=>x.g===g).map(x=>`<label><input type="checkbox" value="${x.id}" ${x.g===cur?"checked":""}> <span>${esc(x.label)}</span></label>`).join("")}</fieldset>`
  });
  $("dlbody").innerHTML = h;
  $("dlmsg").textContent = "";
  const dg = $("dlg");
  $("dlall").onclick = () => dg.querySelectorAll('fieldset input').forEach(x => x.checked = true);
  $("dlnone").onclick = () => dg.querySelectorAll('fieldset input').forEach(x => x.checked = false);
  $("dlx").onclick = () => dg.close();
  $("dlgo").onclick = () => {
    const sel = [...dg.querySelectorAll('fieldset input:checked')].map(x => x.value),
      pick = I.filter(x => sel.indexOf(x.id) >= 0);
    if (!pick.length) {
      $("dlmsg").textContent = "Please tick at least one subsection.";
      return
    }
    const sep = dg.querySelector('input[name=fm]:checked').value === "sep";
    if (!sep) XLSX.writeFile(book(pick.map(x => [x.sheet, x.rows()])), "Chennai_ROTN_Review_Report.xlsx");
    else pick.forEach((x, i) => setTimeout(() => XLSX.writeFile(book([
      [x.sheet, x.rows()]
    ]), x.sheet.replace(/[^A-Za-z0-9]+/g, "_") + ".xlsx"), i * 400));
    $("dlmsg").textContent = sep ? pick.length +
      " files are downloading (allow multiple downloads if your browser asks)." : "Downloaded 1 workbook with " +
      pick.length + " sheet" + (pick.length > 1 ? "s" : "") + ".";
  };
  dg.showModal()
}

/* ============ start ============ */
let cur = "main";
const P = {};

function tab(k) {
  cur = k;
  closePop();
  document.querySelectorAll(".pg").forEach(e => e.hidden = e.id !== "pg-" + k);
  document.querySelectorAll("#tabs button").forEach(b => b.classList.toggle("on", b.dataset.k === k));
  P[k].open();
  scrollTo({
    top: 0
  })
}
fetch("report.json?v=" + Date.now()).then(r => r.json()).then(j => {
  R = j;
  $("asof").textContent = j.asof;
  const t = j.totals;
  const MN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
    "November", "December"
  ];
  const mo = j.month ? MN[+j.month.slice(5) - 1] + " " + j.month.slice(0, 4) : "";
  $("repnote").textContent = "Reviews: review tickets (" + (j.from || "") + " to " + j.asof +
    ") matched with Flipkart / Cloudtail installations closed in any month of the calls report" + (t.fromOther ?
      " (" + t.fromOther + " belong to earlier months)" : "") + ". Denominator: installations closed" + (mo ?
      " in " + mo : "") + " only" + (t.skipped ? " · " + t.skipped +
      " closed tickets of other products (IND, KF etc.) not counted" : "") + ".";
  P.main = mountDash($("pg-main"), "main");
  P.new = mountDash($("pg-new"), "new", j.newRules || j.nr || j);
  P.rej = mountRej($("pg-rej"));
  P.insights = mountInsights($("pg-insights"));
  P.proj = mountProjectedLoss($("pg-proj"));
  P.all = mountDash($("pg-all"), "all");
  const T = [
    ["main", "Review dashboard"],
    ["new", "Dashboard based on new rules"],
    ["rej", "Rejected tickets", j.rejected.length],
    ["insights", "Rejection insights"],
    ["proj", "Projected loss based on rejection rate"],
    ["all", "If all " + t.reviewTickets + " accepted"]
  ];
  $("tabs").innerHTML = T.map(([k, l, n]) =>
    `<button data-k="${k}" role="tab">${l}${n!=null?`<i>${n}</i>`:""}</button>`).join("");
  document.querySelectorAll("#tabs button").forEach(b => b.onclick = () => tab(b.dataset.k));
  tab("main")
}).catch(() => {
  $("asof").textContent = "(report not available yet)";
  $("repnote").textContent = "Report not available yet."
});
$("rdl").onclick = () => {
  if (R) openDl()
};