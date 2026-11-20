function getCore()
{
  const T1 = {
      25: 100,
      50: 150,
      75: 250,
      100: 500
    },
    T2 = {
      25: 150,
      50: 250,
      75: 400,
      100: 750
    };
  /* Two categories, each with its own slab table. Slab % is worked out separately for each category. */
  const CATS = {
    C1: {
      label: "Category 1 – Washing Machine / Cloth Dryer / Microwave",
      short: "Washing Machine, Cloth Dryer & Microwave",
      tbl: 1
    },
    C2: {
      label: "Category 2 – AC / Refrigerator / Dishwasher",
      short: "AC, Refrigerator & Dishwasher",
      tbl: 2
    }
  };
  const ORDER = ["C1", "C2"];
  /* Product column code -> category. Anything else (IND, KF, blank, unknown) is not counted. */
  const PROD = {
    WM: "C1",
    WD: "C1",
    CD: "C1",
    MW: "C1",
    AC: "C2",
    RF: "C2",
    DW: "C2"
  };
  /* E-commerce dealers that count (Dealer Name column) */
  const ECOM = /flipkart india|cloudtail india/i;
  const slabOf = p => p >= 100 ? 100 : p >= 75 ? 75 : p >= 50 ? 50 : p >= 25 ? 25 : 0;
  const nextOf = s => s >= 100 ? 0 : s >= 75 ? 100 : s >= 50 ? 75 : s >= 25 ? 50 : 25;
  const rateOf = (c, s) => s ? (CATS[c].tbl === 1 ? T1 : T2)[s] : 0;
  const clean = s => String(s == null ? "" : s).replace(/[.\s]+/g, " ").trim();
  const title = s => clean(s).toLowerCase().replace(/(^|[\s&\/(-])([a-z])/g, (m, a, b) => a + b.toUpperCase());
  const catOf = p => PROD[String(p == null ? "" : p).trim().toUpperCase()] || null;
  const NEED = ["frcode", "franchise name", "ticket no", "call type", "product", "status", "dealer name", "orgin", "technician name"];
  /* ---- reasons a review ticket is not counted ---- */
  const REASONS = {
    OLD: "Installation call registered before the accepted months (this month and last month only)",
    NF: "Not in the calls file (booked before the file's dates, or not found at all)",
    CANC: "Installation call was Cancelled",
    DEALER: "Closed / soft closure, but dealer is not Flipkart India / Cloudtail India",
    TYPE: "Call type is not Installation",
    STAT: "Status is not Closed / Soft closure",
    PROD: "Product not counted (IND, KF, other)"
  };
  const REASON_ORDER = [REASONS.OLD, REASONS.CANC, REASONS.DEALER, REASONS.STAT, REASONS.TYPE, REASONS.PROD, REASONS.NF];
  const prevMonth = m => {
    let y = +m.slice(0, 4),
      n = +m.slice(5, 7) - 1;
    if (n < 1) {
      n = 12;
      y--
    }
    return y + "-" + String(n).padStart(2, "0")
  };

  function stat(c, closed, reviews, wrong) {
    const pct = closed ? reviews / closed * 100 : (reviews ? 100 : 0),
      slab = slabOf(pct),
      rate = rateOf(c, slab),
      next = nextOf(slab);
    const need = next ? Math.max(1, Math.ceil(next * closed / 100) - reviews) : 0;
    return {
      closed,
      reviews,
      wrong,
      pct: Math.round(pct * 100) / 100,
      slab,
      rate,
      earn: rate * reviews,
      next,
      need,
      nextRate: next ? rateOf(c, next) : 0,
      nextEarn: next ? rateOf(c, next) * (reviews + need) : 0
    }
  }
  const DATEPRI = [/regist.*date|date.*regist/, /call.*date|date.*call/, /creat|logged/, /^date$/, /date/],
    NOTDATE = /clos|complet|resolv|alloc|appoint|schedul|visit|purchase|warrant|due|tat|aging|age/;

  function pickDateCol(names) {
    for (const re of DATEPRI) {
      const j = names.findIndex(n => re.test(n) && !NOTDATE.test(n));
      if (j >= 0) return j
    }
    return -1
  }

  function ymd(v, dmy) {
    if (v == null || v === "") return null;
    if (typeof v === "number") {
      if (v < 20000 || v > 80000) return null;
      const d = new Date(Math.round((v - 25569) * 864e5));
      return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]
    }
    if (v instanceof Date) return [v.getFullYear(), v.getMonth() + 1, v.getDate()];
    const m = String(v).match(/(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})/);
    if (!m) return null;
    let a = +m[1],
      b = +m[2],
      c = +m[3],
      r;
    if (m[1].length === 4) r = [a, b, c];
    else {
      if (c < 100) c += 2000;
      r = dmy ? [c, b, a] : [c, a, b]
    }
    return r[1] >= 1 && r[1] <= 12 ? r : null
  }

  function pickCalls(aoa) {
    let h = -1,
      range = "";
    for (let i = 0; i < Math.min(aoa.length, 15); i++) {
      const c = String(aoa[i][0] || "").trim();
      if (/^date range/i.test(c)) range = c.replace(/^date range\s*:?\s*/i, "");
      if (c.toLowerCase() === "branch name") {
        h = i;
        break
      }
    }
    if (h < 0) throw new Error("Could not find the header row (Branch Name …). Is this the Calls Registered report?");
    const names = aoa[h].map(n => String(n).trim().toLowerCase()),
      ix = {};
    names.forEach((n, j) => {
      ix[n] = j
    });
    const miss = NEED.filter(n => ix[n] === undefined);
    if (miss.length) throw new Error("Missing columns: " + miss.join(", "));
    const dc = pickDateCol(names),
      raw = [];
    for (let i = h + 1; i < aoa.length; i++) {
      const r = aoa[i];
      if (!r || !r[ix["ticket no"]]) continue;
      raw.push(NEED.map(n => r[ix[n]]).concat([dc >= 0 ? r[dc] : null]))
    }
    /* month of every call (YYYY-MM): decide dd/mm or mm/dd from the data itself */
    let dmy = false;
    if (dc >= 0) raw.forEach(r => {
      const m = typeof r[9] === "string" && r[9].match(/^\s*(\d{1,2})[\/\-.](\d{1,2})[\/\-.]\d{2,4}/);
      if (m && +m[1] > 12) dmy = true
    });
    const months = {};
    let ok = 0;
    const p2 = n => String(n).padStart(2, "0");
    raw.forEach(r => {
      const d = dc >= 0 ? ymd(r[9], dmy) : null;
      r[9] = d ? d[0] + "-" + p2(d[1]) : null;
      r[10] = d ? d[0] + "-" + p2(d[1]) + "-" + p2(d[2]) : "";
      if (r[9]) {
        ok++;
        months[r[9]] = (months[r[9]] || 0) + 1
      }
    });
    return {
      rows: raw,
      range: range.replace(/\s+12:00:00 AM/g, ""),
      dateCol: dc >= 0 ? String(aoa[h][dc]).trim() : "",
      months,
      dateOk: dc >= 0 && raw.length > 0 && ok / raw.length >= 0.5
    }
  }

  function buildReport(rows, rev, opts) {
    opts = opts || {};
    const closedSt = opts.softClosed === false ? ["Closed"] : ["Closed", "Resolved (soft closure)"],
      lc = closedSt.map(s => s.toLowerCase());
    const F = {},
      seen = {};
    let closed = 0,
      reviews = 0,
      skipped = 0;
    const used = {},
      first = {};
    const mon = opts.month || "";
    /* NEW RULES: opts.callMonths = months (YYYY-MM) in which the e-com installation call must have been registered (report month + month before).
       A review ticket whose matching closed installation call is from an older month is NOT accepted. The upload date of the review does not matter. */
    const cm = opts.callMonths || null,
      old = {};
    const reviewTickets = Object.keys(rev).length;
    const inFile = {};
    let exProd = 0,
      fromOther = 0;
    const isE = r => String(r[3]).trim().toUpperCase() === "INSTALLATION CALL" && ECOM.test(String(r[6])) && lc.indexOf(String(r[5]).trim().toLowerCase()) >= 0;
    /* Denominator: e-com installation tickets closed in the chosen month only.
       Numerator: every closed e-com installation ticket of ANY month in the calls file that is in the review data. */
    rows.forEach(r => {
      const [fc, fn, tk, ct, pr, st, dl, og, tn] = r;
      const t = String(tk).trim();
      const known = !!rev[t];
      if (known) {
        inFile[t] = 1;
        if (!first[t] || (!isE(first[t]) && isE(r))) first[t] = r
      }
      if (!isE(r) || seen[t]) return;
      seen[t] = 1;
      const isRev = known && (!cm || cm.indexOf(r[9]) >= 0);
      if (known && !isRev) old[t] = 1;
      const inM = !mon || r[9] === mon;
      if (!inM && !isRev) return;
      const c = catOf(pr);
      if (!c) {
        if (isRev) exProd++;
        else skipped++;
        return
      }
      const f = F[fc] || (F[fc] = {
        code: String(fc),
        name: title(fn),
        techs: {}
      });
      const key = clean(tn).toUpperCase() || "(NO TECHNICIAN)";
      const te = f.techs[key] || (f.techs[key] = {
        name: title(tn) || "(No technician)",
        cats: {}
      });
      const ce = te.cats[c] || (te.cats[c] = {
        closed: 0,
        reviews: 0,
        extra: 0
      });
      if (inM) {
        ce.closed++;
        closed++
      }
      if (isRev) {
        ce.reviews++;
        reviews++;
        used[t] = 1;
        if (!inM) fromOther++
      }
    });
    /* Scenario "all uploaded tickets accepted": every review ticket that was not counted (cancelled call, other dealer, ...) is added
       to the numerator of its franchise / technician / category. The denominator stays the same. Tickets with no call row cannot be placed. */
    let unattr = 0;
    Object.keys(rev).forEach(t => {
      if (used[t]) return;
      const r = first[t],
        c = r && catOf(r[4]);
      if (!c) {
        unattr++;
        return
      }
      const fc = r[0],
        f = F[fc] || (F[fc] = {
          code: String(fc),
          name: title(r[1]),
          techs: {}
        }),
        key = clean(r[8]).toUpperCase() || "(NO TECHNICIAN)";
      const te = f.techs[key] || (f.techs[key] = {
          name: title(r[8]) || "(No technician)",
          cats: {}
        }),
        ce = te.cats[c] || (te.cats[c] = {
          closed: 0,
          reviews: 0,
          extra: 0
        });
      ce.extra++
    });
    const pc = (rv, cl) => cl ? Math.round(rv / cl * 10000) / 100 : (rv ? 100 : 0);
    let earn = 0,
      aReviews = 0,
      aEarn = 0;
    const fr = Object.values(F).map(f => {
        const techs = Object.values(f.techs).map(te => {
            const cats = {};
            let cl = 0,
              rv = 0,
              er = 0,
              ar = 0,
              ae = 0,
              ex = 0;
            ORDER.forEach(c => {
              const x = te.cats[c];
              if (x) {
                const s = stat(c, x.closed, x.reviews, 0),
                  a = stat(c, x.closed, x.reviews + x.extra, 0);
                cats[c] = Object.assign(s, {
                  extra: x.extra,
                  a
                });
                cl += x.closed;
                rv += x.reviews;
                er += s.earn;
                ar += a.reviews;
                ae += a.earn;
                ex += x.extra
              }
            });
            return {
              name: te.name,
              closed: cl,
              reviews: rv,
              pct: pc(rv, cl),
              earn: er,
              extra: ex,
              aReviews: ar,
              aPct: pc(ar, cl),
              aEarn: ae,
              gain: ae - er,
              cats
            }
          })
          .sort((a, b) => b.earn - a.earn || b.pct - a.pct || b.closed - a.closed || a.name.localeCompare(b.name));
        const S = k => techs.reduce((a, t) => a + t[k], 0),
          cl = S("closed"),
          rv = S("reviews"),
          er = S("earn"),
          ar = S("aReviews"),
          ae = S("aEarn");
        earn += er;
        aReviews += ar;
        aEarn += ae;
        return {
          code: f.code,
          name: f.name,
          closed: cl,
          reviews: rv,
          pct: pc(rv, cl),
          earn: er,
          extra: S("extra"),
          aReviews: ar,
          aPct: pc(ar, cl),
          aEarn: ae,
          gain: ae - er,
          techs
        }
      })
      .sort((a, b) => a.name.localeCompare(b.name));
    /* review tickets that were not counted, with the reason */
    const why = r => {
      if (!r) return [REASONS.NF, ""];
      const [fc, fn, tk, ct, pr, st, dl, og] = r, stl = String(st).trim().toLowerCase(), de = ECOM.test(String(dl));
      if (String(ct).trim().toUpperCase() !== "INSTALLATION CALL") return [REASONS.TYPE, String(ct)];
      if (stl === "cancelled") return [REASONS.CANC, de ? "" : "Dealer is also not Flipkart India / Cloudtail India"];
      if (lc.indexOf(stl) < 0) return [REASONS.STAT, String(st)];
      if (!de) return [REASONS.DEALER, ""];
      if (!catOf(pr)) return [REASONS.PROD, String(pr)];
      return ["Not counted", ""]
    };
    const info = (t, r) => ({
      t,
      b: rev[t].b || "",
      d: rev[t].d || "",
      fc: r ? String(r[0]) : "",
      fn: r ? title(r[1]) : "",
      tn: r ? title(r[8]) : "",
      st: r ? String(r[5]) : "",
      dl: r ? String(r[6]) : "",
      og: r ? String(r[7]) : "",
      pr: r ? String(r[4]) : "",
      ct: r ? String(r[3]) : "",
      cb: r ? (r[10] || "") : ""
    });
    const MNs = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
      mlab = m => MNs[+m.slice(5, 7) - 1] + " " + m.slice(0, 4);
    const rejected = Object.keys(rev).filter(t => !used[t]).map(t => {
      const r = first[t];
      let w = why(r);
      if (old[t] && w[0] === "Not counted") w = [REASONS.OLD, "Call registered " + (r[10] || r[9] || "?") + " · accepted only " + cm.slice().reverse().map(mlab).join(" and ")];
      return Object.assign(info(t, r), {
        why: w[0],
        note: w[1]
      })
    });
    const unmatched = rejected.length,
      notFound = rejected.filter(x => x.why === REASONS.NF).length,
      stale = rejected.filter(x => x.why === REASONS.OLD).length;
    return {
      version: 4,
      range: opts.range || "",
      asof: opts.asof || "",
      from: opts.from || "",
      month: mon,
      softClosed: opts.softClosed !== false,
      totals: {
        closed,
        reviews,
        fromOther,
        reviewTickets,
        unmatched,
        notFound,
        exProd,
        skipped,
        earn,
        pct: pc(reviews, closed),
        unattr,
        stale,
        aReviews,
        aPct: pc(aReviews, closed),
        aEarn,
        gain: aEarn - earn
      },
      franchises: fr,
      rejected,
      window: cm ? {
        months: cm,
        basis: "call"
      } : null
    }
  }
  /* Current rules report + the same report under the NEW RULES in .nr (accepted only if the matching closed e-com installation call was registered in the report month or the month before).
     The "if all uploaded tickets were accepted" figures (aReviews, aEarn, gain) of .nr are measured against the new rules. */
  function buildBoth(rows, rev, opts) {
    opts = opts || {};
    const rep = buildReport(rows, rev, opts),
      m = opts.month || String(opts.asof || "").slice(0, 7);
    if (m && rows.some(r => r[9])) rep.nr = buildReport(rows, rev, Object.assign({}, opts, {
      callMonths: [m, prevMonth(m)]
    }));
    return rep
  }
  /* Flat Excel rows: one row per technician per category. withAll = also add the "if all uploaded tickets are accepted" columns */
  function techRows(f, only, withAll) {
    const out = [];
    f.techs.forEach(t => ORDER.forEach(c => {
      if (only && only !== "all" && only !== c) return;
      const x = t.cats[c];
      if (!x) return;
      const o = {
        "Franchise": f.name,
        "Franchise code": f.code,
        "Technician": t.name,
        "Category": CATS[c].label,
        "Closed (category)": x.closed,
        "Reviews (category)": x.reviews,
        "Review % (category)": x.pct,
        "Slab %": x.slab,
        "₹ per review": x.rate,
        "Earned ₹": x.earn,
        "Next slab %": x.next || "Top slab",
        "Reviews needed for next slab": x.next ? x.need : 0,
        "Earning at next slab ₹": x.next ? x.nextEarn : x.earn
      };
      if (withAll) {
        o["Rejected tickets accepted"] = x.extra;
        o["Reviews if all accepted"] = x.reviews + x.extra;
        o["Slab % if all accepted"] = x.a.slab;
        o["Earned ₹ if all accepted"] = x.a.earn;
        o["Extra ₹ if all accepted"] = x.a.earn - x.earn
      }
      o["Technician total earned ₹"] = t.earn;
      if (withAll) o["Technician total ₹ if all accepted"] = t.aEarn;
      out.push(o)
    }));
    return out
  }
  const api = {
    CATS,
    ORDER,
    PROD,
    T1,
    T2,
    REASONS,
    REASON_ORDER,
    slabOf,
    nextOf,
    rateOf,
    stat,
    prevMonth,
    pickCalls,
    buildReport,
    buildBoth,
    techRows
  };

  return api;
};