/* 财报跟踪平台 · 前端逻辑（纯 JS，无框架；数据来自 data/data.js 的 window.SITE_DATA） */
(function () {
  "use strict";

  /* 自诊断：脚本错误浮条（用户截图即可定位问题） */
  window.addEventListener("error", function (e) {
    var b = document.getElementById("js-err-banner");
    if (!b) {
      b = document.createElement("div");
      b.id = "js-err-banner";
      b.style.cssText = "position:fixed;bottom:0;left:0;right:0;background:#5d1f1f;color:#ffd7d7;" +
        "padding:8px 14px;font-size:13px;z-index:99999;font-family:monospace";
      document.body.appendChild(b);
    }
    b.textContent = "页面脚本错误：" + e.message + "（" + String(e.filename || "").split("/").pop() +
      ":" + e.lineno + "）— 请截图发给维护者";
  });

  /* ================= 密码门（本地防窥门槛，非安全加密） ================= */
  // SHA-256("访问密码")，明文不出现在界面与源码中（README 除外）
  var GATE_HASH = "dfc1d541e6dbbc1f24d98dde8da2f19bd6fc57565ff43ff04a012a12958966ca";
  var GATE_KEY = "yafco_gate_ok";

  function sha256Hex(text) {
    var enc = new TextEncoder().encode(text);
    return crypto.subtle.digest("SHA-256", enc).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) {
        return ("0" + b.toString(16)).slice(-2);
      }).join("");
    });
  }

  function initGate(onPass) {
    var gate = document.getElementById("gate");
    var input = document.getElementById("gate-input");
    var btn = document.getElementById("gate-btn");
    var err = document.getElementById("gate-err");
    if (sessionStorage.getItem(GATE_KEY) === "1") {
      gate.style.display = "none";
      onPass();
      return;
    }
    function tryPass() {
      var v = input.value;
      if (!v) { err.textContent = "请输入密码"; return; }
      sha256Hex(v).then(function (h) {
        if (h === GATE_HASH) {
          sessionStorage.setItem(GATE_KEY, "1");
          gate.style.display = "none";
          onPass();
        } else {
          err.textContent = "密码错误，请重试";
          input.value = "";
          input.focus();
        }
      });
    }
    btn.addEventListener("click", tryPass);
    input.addEventListener("keydown", function (e) { if (e.key === "Enter") tryPass(); });
    input.focus();
  }

  /* ================= 工具 ================= */
  function $(sel, root) { return (root || document).querySelector(sel); }
  // 区块折叠：点击 section-head 展开/收起，箭头随动；defaultCollapsed=true 时默认收起
  function makeCollapsible(sectionEl, defaultCollapsed) {
    var head = sectionEl.querySelector(".section-head");
    if (!head) return;
    head.classList.add("collapsible");
    head.title = "点击折叠 / 展开";
    if (defaultCollapsed) sectionEl.classList.add("collapsed");
    head.addEventListener("click", function (e) {
      if (e.target.closest("button, a, input, .view-toggle")) return;  // 头部内按钮不触发折叠
      sectionEl.classList.toggle("collapsed");
    });
  }
  function el(tag, cls, html) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (html !== undefined && html !== null) e.innerHTML = html;
    return e;
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function fmtNum(v, digits) {
    if (v === null || v === undefined || isNaN(v)) return "—";
    return Number(v).toLocaleString("zh-CN", { maximumFractionDigits: digits === undefined ? 2 : digits });
  }
  function fmtPct(v) {
    if (v === null || v === undefined || isNaN(v)) return '<span class="flat">—</span>';
    var pct = (v * 100);
    var cls = pct > 0.0001 ? "up" : (pct < -0.0001 ? "down" : "flat");
    var sign = pct > 0 ? "+" : "";
    return '<span class="' + cls + '">' + sign + pct.toFixed(1) + "%</span>";
  }
  // 指引进度状态标签：超出=红（涨）、不及=绿（跌）、符合/季节性=灰
  function statusBadge(status) {
    if (!status) return '<span class="flat">—</span>';
    var cls = status === "超出" ? "up" : (status === "不及" ? "down" : "flat");
    return '<span class="' + cls + '">' + esc(status) + "</span>";
  }
  function quarterKey(q) {
    var m = /^(\d{4})Q([1-4])$/.exec(q);
    return m ? (+m[1]) * 10 + (+m[2]) : 0;
  }
  function periodKey(p) {
    if (/^\d{4}$/.test(p)) return (+p) * 10;
    return quarterKey(p);
  }
  // 从指引字符串解析数值（区间取中值）：'109-110' -> 109.5, '32000' -> 32000, '5,800–6,100' -> 5950
  function parseGuide(s) {
    if (!s) return null;
    var m = String(s).match(/(\d[\d,]*(?:\.\d+)?)\s*[-~—–]\s*(\d[\d,]*(?:\.\d+)?)/);
    if (m) return (parseFloat(m[1].replace(/,/g, "")) + parseFloat(m[2].replace(/,/g, ""))) / 2;
    m = String(s).match(/(\d[\d,]*(?:\.\d+)?)/);
    return m ? parseFloat(m[1].replace(/,/g, "")) : null;
  }
  function latestPeriodWithData(data, quarters, years) {
    var ps = (quarters || []).concat(years || []);
    var best = null;
    ps.forEach(function (p) {
      if (data[p] !== null && data[p] !== undefined) {
        if (best === null || periodKey(p) > periodKey(best)) best = p;
      }
    });
    return best;
  }
  function hasEcharts() { return typeof window.echarts !== "undefined"; }
  function darkChartBase() {
    return {
      backgroundColor: "transparent",
      textStyle: { color: "#6d6248", fontFamily: "Microsoft YaHei, sans-serif" },
      tooltip: {
        trigger: "axis",
        backgroundColor: "#fdfaf3", borderColor: "#ddd5c4",
        textStyle: { color: "#2e2a22", fontSize: 12 }
      },
      grid: { left: 60, right: 60, top: 50, bottom: 40 }
    };
  }
  var PALETTE = ["#b8912f", "#4a6b8a", "#c0453e", "#2a9e6e", "#7a5f9e", "#b06a3b", "#2aa89e",
    "#a04a6e", "#7a9a3a", "#5a6fb0", "#c9a52e", "#6d6248", "#a06a3a", "#3a8a9e", "#7a7ac0",
    "#c05a5a", "#4a9a4a", "#a88a3a", "#4a8ab0", "#9a7ab0", "#7a9a3a"];

  /* ================= 区块：产量（表格 + 图表） ================= */
  var viewState = {}; // sectionDomId -> "quarter" | "year"
  var filterState = {}; // sectionDomId -> { per:{期间:true=隐藏}, comp:{公司key:true=隐藏} }

  function renderProductionSection(container, sec, commodityKey, secKey, defaultView) {
    var domId = commodityKey + "-" + secKey;
    viewState[domId] = viewState[domId] || defaultView || "quarter";

    var section = el("div", "section");
    var head = el("div", "section-head");
    head.appendChild(el("div", "section-title", esc(sec.title)));
    head.appendChild(el("div", "section-sub", "单位：" + esc(sec.unit)));
    if (sec.companies.some(function (c) { return c.est_q; })) {
      head.appendChild(el("div", "section-sub", "斜体/琥珀色 = 拟合季度值（悬停看方法）"));
    }
    var toggle = el("div", "view-toggle");
    var bq = el("button", viewState[domId] === "quarter" ? "active" : "", "季度视图");
    var by = el("button", viewState[domId] === "year" ? "active" : "", "年度视图");
    toggle.appendChild(bq); toggle.appendChild(by);
    head.appendChild(toggle);
    section.appendChild(head);

    var chartBox = el("div", "panel");
    var chartDiv = el("div", "chart");
    chartDiv.id = "chart-" + domId;
    chartBox.appendChild(chartDiv);
    section.appendChild(chartBox);

    var tablePanel = el("div", "panel");
    tablePanel.style.marginTop = "14px";
    var filterBar = el("div", "tbl-filter");
    tablePanel.appendChild(filterBar);
    var wrap = el("div", "table-wrap freeze");
    tablePanel.appendChild(wrap);
    section.appendChild(tablePanel);
    container.appendChild(section);

    var chartInst = null;
    var fState = filterState[domId] = filterState[domId] || { per: {}, comp: {}, collapsed: {}, sort: null };

    function compKey(c) { return c.name + (c.project ? "·" + c.project : ""); }
    function prevPeriod(p) {
      var m = /^(\d{4})Q(\d)$/.exec(p);
      if (m) return (parseInt(m[1], 10) - 1) + "Q" + m[2];
      if (/^\d{4}$/.test(p)) return String(parseInt(p, 10) - 1);
      return null;
    }
    // 按一级公司分组：优先 c.group（母公司映射，如 Vedanta 系/Boliden 两厂），否则按 c.name；
    // 多行同组 → 父行（小计，可折叠）+ 子行（项目）
    function buildGroups(comps) {
      var groups = [], byName = {};
      comps.forEach(function (c) {
        var key = c.group || c.name;
        var g = byName[key];
        if (!g) { g = byName[key] = { name: key, items: [] }; groups.push(g); }
        g.items.push(c);
      });
      return groups;
    }
    function groupSum(items, p) {
      var s = 0, seen = false;
      items.forEach(function (c) { var x = c.data[p]; if (x !== null && x !== undefined) { s += x; seen = true; } });
      return seen ? s : null;
    }
    function cellVal(c, p) { var v = c.data[p]; return (v === null || v === undefined) ? null : v; }
    function cmpVal(a, b) {  // null 一律排最后
      if (a === null && b === null) return 0;
      if (a === null) return 1;
      if (b === null) return -1;
      return a < b ? -1 : a > b ? 1 : 0;
    }

    function periods() {
      return viewState[domId] === "quarter" ? sec.quarters : sec.years;
    }

    function renderTable() {
      var ps = periods().filter(function (p) { return !fState.per[p]; });
      var comps = sec.companies.filter(function (c) { return !fState.comp[compKey(c)]; });
      var compFiltered = comps.length !== sec.companies.length;
      var groups = buildGroups(comps);
      // 列排序：点期间表头按该列值排序（组按小计、组内按子行值；null 排最后）；再点切换升降；点「公司」列恢复默认
      var sort = fState.sort;
      if (sort && sort.p) {
        groups.forEach(function (g) {
          g.items.sort(function (a, b) { return cmpVal(cellVal(a, sort.p), cellVal(b, sort.p)) * sort.dir; });
          g._sum = groupSum(g.items, sort.p);
        });
        groups.sort(function (a, b) { return cmpVal(a._sum, b._sum) * sort.dir; });
      }
      // 真·分组按「全量公司」判定（筛选只剩一行时不降级为独立行，避免归属误读）
      var trueGroupKeys = {};
      buildGroups(sec.companies).forEach(function (g) { if (g.items.length > 1) trueGroupKeys[g.name] = true; });
      var html = '<table class="data-table"><thead><tr><th class="sortable" data-sort="" title="点击恢复默认顺序">公司</th><th>国家/地区</th>';
      ps.forEach(function (p) {
        var arrow = sort && sort.p === p ? (sort.dir < 0 ? " ▼" : " ▲") : "";
        html += '<th class="sortable" data-sort="' + esc(p) + '" title="点击按本列排序">' + esc(p) + arrow + "</th>";
      });
      html += "<th>最新同比</th></tr></thead><tbody>";
      groups.forEach(function (g) {
        var isGroup = !!trueGroupKeys[g.name];
        var collapsed = isGroup && fState.collapsed && fState.collapsed[g.name];
        if (isGroup) {  // 一级公司父行：小计 + 折叠开关
          html += '<tr class="grp-row" data-grp="' + esc(g.name) + '" title="点击折叠/展开"><td>' +
            '<span class="grp-toggle">' + (collapsed ? "▸" : "▾") + "</span> <b>" + esc(g.name) + "</b>" +
            " <span class='grp-n'>" + g.items.length + " 项目</span></td><td></td>";
          ps.forEach(function (p) {
            var s = groupSum(g.items, p);
            html += s === null ? '<td class="na">—</td>' : '<td class="grp-sum">' + fmtNum(s) + "</td>";
          });
          var lp = null;
          for (var i = ps.length - 1; i >= 0; i--) { if (groupSum(g.items, ps[i]) !== null) { lp = ps[i]; break; } }
          var gy;
          if (lp) {
            var gpp = prevPeriod(lp), gsv = groupSum(g.items, lp), gsp = gpp ? groupSum(g.items, gpp) : null;
            if (gsv !== null && gsp) gy = (gsv - gsp) / gsp;
          }
          html += "<td>" + (gy !== undefined ? fmtPct(gy) : '<span class="flat">—</span>') + "</td></tr>";
        }
        if (collapsed) return;
        g.items.forEach(function (c) {
          var est = c.est ? '<sup style="color:#b8912f" title="含估算/推算：' + esc(c.est_note || "") + '">†</sup>' : "";
          var label = isGroup && c.project
            ? esc(c.project) + est
            : esc(c.name) + est + (c.project ? ' <span style="color:#5b6879">· ' + esc(c.project) + "</span>" : "");
          if (!isGroup) label += ' <span class="solo-badge" title="独立公司：不属于任何集团分组">独立</span>';
          var nameTitle = c.est && c.est_note ? ' title="含估算/推算：' + esc(c.est_note) + '"' : "";
          html += "<tr><td" + (isGroup ? ' class="child-cell"' : "") + nameTitle + ">" + label + "</td><td>" + esc(c.country || "—") + "</td>";
          ps.forEach(function (p) {
            var v = c.data[p];
            var isEstQ = c.est_q && c.est_q[p];
            var title = "";
            if (isEstQ && c.est_q_note && c.est_q_note[p]) title = ' title="' + esc(c.est_q_note[p]) + '"';
            else if (c.reason && v !== null && v !== undefined) title = ' title="变化原因：' + esc(c.reason) + '"';
            html += v === null || v === undefined
              ? '<td class="na">—</td>'
              : '<td class="' + (isEstQ ? "estq" : "") + '"' + title + ">" + fmtNum(v) + "</td>";
          });
          var lp2 = latestPeriodWithData(c.data, ps, []);
          html += "<td>" + (lp2 && c.yoy[lp2] !== undefined ? fmtPct(c.yoy[lp2]) : '<span class="flat">—</span>') + "</td></tr>";
        });
      });
      // 总计行：默认用同口径预计算合计；公司筛选后按可见公司重算（简单同比）
      html += '<tr class="total-row"><td>合计' +
        (compFiltered ? ' <small style="font-weight:400">（可见公司）</small>' : "") + "</td><td></td>";
      ps.forEach(function (p) {
        var v, y;
        if (!compFiltered) { v = sec.total.data[p]; y = sec.total.yoy[p]; }
        else {
          v = groupSum(comps, p);
          var pp = prevPeriod(p), sp = pp ? groupSum(comps, pp) : null;
          y = (v !== null && sp) ? (v - sp) / sp : undefined;
        }
        html += v === null || v === undefined
          ? '<td class="na">—</td>'
          : "<td>" + fmtNum(v) + (y !== undefined ? " <small>" + fmtPct(y) + "</small>" : "") + "</td>";
      });
      html += "<td></td></tr></tbody></table>";
      wrap.innerHTML = html;
      // 绑定：父行折叠/展开 + 表头排序
      wrap.querySelectorAll(".grp-row").forEach(function (tr) {
        tr.addEventListener("click", function () {
          var n = tr.getAttribute("data-grp");
          fState.collapsed[n] = !fState.collapsed[n];
          renderTable();
        });
      });
      wrap.querySelectorAll("th.sortable").forEach(function (th) {
        th.addEventListener("click", function () {
          var sp = th.getAttribute("data-sort");
          if (!sp) fState.sort = null;
          else if (fState.sort && fState.sort.p === sp) fState.sort.dir = -fState.sort.dir;
          else fState.sort = { p: sp, dir: -1 };  // 首次点击默认降序（大的在前）
          renderTable();
        });
      });
    }

    // 表格筛选工具条：期间（季度视图按年分组，可整年折叠）+ 公司两个多选下拉；
    // 只作用于表格，图表系列用 ECharts 图例单独开关
    function renderFilterBar() {
      var ps = periods();
      var groups = [];
      if (viewState[domId] === "quarter") {
        var byYear = {};
        ps.forEach(function (p) { var y = p.slice(0, 4); (byYear[y] = byYear[y] || []).push(p); });
        Object.keys(byYear).sort().forEach(function (y) { groups.push({ label: y + " 年", items: byYear[y] }); });
      } else {
        ps.forEach(function (p) { groups.push({ label: p, items: [p] }); });
      }
      var compGroups = buildGroups(sec.companies);  // 公司筛选也按一级公司分组
      filterBar.innerHTML =
        '<details class="tf"><summary>期间筛选<b class="tf-n"></b></summary><div class="tf-body">' +
        groups.map(function (g, gi) {
          var inner = g.items.length > 1
            ? g.items.map(function (p) {
                return '<label class="tf-item"><input type="checkbox" data-per="' + esc(p) + '"> ' + esc(p) + "</label>";
              }).join("")
            : "";
          return '<div class="tf-group"><label class="tf-master"><input type="checkbox" data-master="' + gi + '"> ' +
            esc(g.label) + "</label>" + inner + "</div>";
        }).join("") +
        '</div></details>' +
        '<details class="tf"><summary>公司筛选<b class="tf-n"></b></summary><div class="tf-body">' +
        compGroups.map(function (g, gi) {
          if (g.items.length === 1) {
            return '<label class="tf-item"><input type="checkbox" data-comp="' + esc(compKey(g.items[0])) + '"> ' +
              esc(compKey(g.items[0])) + "</label>";
          }
          return '<div class="tf-group"><label class="tf-master"><input type="checkbox" data-cmaster="' + gi + '"> ' +
            esc(g.name) + "</label>" +
            g.items.map(function (c) {
              return '<label class="tf-item"><input type="checkbox" data-comp="' + esc(compKey(c)) + '"> ' +
                esc(c.project || c.name) + "</label>";
            }).join("") + "</div>";
        }).join("") +
        '</div></details>' +
        '<button class="tf-reset" type="button">重置筛选</button><span class="tf-hint"></span>';

      function refresh() {  // 只同步勾选态/角标/提示，不重建 DOM（保持下拉展开状态）
        filterBar.querySelectorAll("[data-per]").forEach(function (box) {
          box.checked = !fState.per[box.getAttribute("data-per")];
        });
        filterBar.querySelectorAll("[data-comp]").forEach(function (box) {
          box.checked = !fState.comp[box.getAttribute("data-comp")];
        });
        filterBar.querySelectorAll("[data-master]").forEach(function (box) {
          var g = groups[+box.getAttribute("data-master")];
          var off = g.items.filter(function (p) { return fState.per[p]; }).length;
          box.checked = off === 0;
          box.indeterminate = off > 0 && off < g.items.length;
        });
        filterBar.querySelectorAll("[data-cmaster]").forEach(function (box) {
          var g = compGroups[+box.getAttribute("data-cmaster")];
          var off = g.items.filter(function (c) { return fState.comp[compKey(c)]; }).length;
          box.checked = off === 0;
          box.indeterminate = off > 0 && off < g.items.length;
        });
        var np = ps.filter(function (p) { return fState.per[p]; }).length;
        var nc = sec.companies.filter(function (c) { return fState.comp[compKey(c)]; }).length;
        var sns = filterBar.querySelectorAll(".tf-n");
        sns[0].textContent = np ? " 隐 " + np : "";
        sns[1].textContent = nc ? " 隐 " + nc : "";
        filterBar.querySelector(".tf-hint").textContent = (np || nc)
          ? "已隐藏 " + np + " 个期间、" + nc + " 家公司" +
            (nc ? "；合计按可见公司重算（简单同比）" : "") + "。图表用图例单独开关"
          : "可按年/季度、按公司折叠隐藏；图表系列用图例开关";
      }
      filterBar.querySelectorAll("[data-per]").forEach(function (box) {
        box.addEventListener("change", function () {
          var p = box.getAttribute("data-per");
          if (box.checked) delete fState.per[p]; else fState.per[p] = true;
          refresh(); renderTable();
        });
      });
      filterBar.querySelectorAll("[data-master]").forEach(function (box) {
        box.addEventListener("change", function () {
          var g = groups[+box.getAttribute("data-master")];
          g.items.forEach(function (p) { if (box.checked) delete fState.per[p]; else fState.per[p] = true; });
          refresh(); renderTable();
        });
      });
      filterBar.querySelectorAll("[data-cmaster]").forEach(function (box) {
        box.addEventListener("change", function () {
          var g = compGroups[+box.getAttribute("data-cmaster")];
          g.items.forEach(function (c) {
            var k = compKey(c);
            if (box.checked) delete fState.comp[k]; else fState.comp[k] = true;
          });
          refresh(); renderTable();
        });
      });
      filterBar.querySelectorAll("[data-comp]").forEach(function (box) {
        box.addEventListener("change", function () {
          var k = box.getAttribute("data-comp");
          if (box.checked) delete fState.comp[k]; else fState.comp[k] = true;
          refresh(); renderTable();
        });
      });
      filterBar.querySelector(".tf-reset").addEventListener("click", function () {
        fState.per = {}; fState.comp = {}; refresh(); renderTable();
      });
      refresh();
    }

    function renderChart() {
      if (!hasEcharts()) {
        chartDiv.outerHTML = '<div class="chart-fallback">图表加载失败（ECharts CDN 不可用），数据请见下表</div>';
        return;
      }
      if (!chartInst) chartInst = echarts.init(chartDiv, null, { renderer: "svg" });  // svg 渲染：规避部分机器 GPU canvas 加速故障导致的空白图
      var ps = periods();
      var estNotes = {};
      var estQBySeries = {};
      var series = sec.companies.map(function (c, i) {
        var sName = (c.project ? c.name + "·" + c.project : c.name) + (c.est ? " †" : "");
        if (c.est && c.est_note) estNotes[sName] = c.est_note;
        if (c.est_q) estQBySeries[sName] = c.est_q;
        return {
          name: sName,
          type: "bar", stack: "total", barMaxWidth: 34,
          itemStyle: { color: PALETTE[i % PALETTE.length] },
          emphasis: { focus: "series" },
          data: ps.map(function (p) {
            var v = c.data[p] === undefined ? null : c.data[p];
            // 拟合季度：低透明度 + 琥珀色虚线描边区分
            if (v !== null && c.est_q && c.est_q[p]) {
              return { value: v, itemStyle: { opacity: 0.45, borderColor: "#b8912f", borderWidth: 1, borderType: "dashed" } };
            }
            return v;
          })
        };
      });
      series.push({
        name: "合计同比", type: "line", yAxisIndex: 1,
        symbol: "circle", symbolSize: 6,
        lineStyle: { color: "#b8912f", width: 2, type: "dashed" },
        itemStyle: { color: "#b8912f" },
        data: ps.map(function (p) {
          var y = sec.total.yoy[p];
          return y === undefined || y === null ? null : +(y * 100).toFixed(2);
        })
      });
      var opt = Object.assign(darkChartBase(), {
        legend: { type: "scroll", top: 4, textStyle: { color: "#6d6248", fontSize: 11 } },
        tooltip: {
          trigger: "axis",
          backgroundColor: "#fdfaf3", borderColor: "#ddd5c4",
          textStyle: { color: "#2e2a22", fontSize: 12 },
          formatter: function (params) {
            var s = esc(params[0].axisValue);
            var notes = [];
            params.forEach(function (p) {
              if (p.value === null || p.value === undefined) return;
              s += "<br/>" + p.marker + esc(p.seriesName) + "：" + fmtNum(p.value) + (p.seriesName === "合计同比" ? "%" : "");
              if (estQBySeries[p.seriesName] && estQBySeries[p.seriesName][p.axisValue]) {
                s += " <span style='color:#b8912f'>(拟合)</span>";
              }
              if (estNotes[p.seriesName] && notes.indexOf(p.seriesName) < 0) notes.push(p.seriesName);
            });
            if (notes.length) {
              var first = estNotes[notes[0]] || "";
              s += "<br/><span style='color:#b8912f;font-size:11px'>† 含估算/推算：" +
                esc(first.length > 160 ? first.slice(0, 160) + "…" : first) +
                (notes.length > 1 ? "（另有 " + (notes.length - 1) + " 条，详见表格公司名悬停）" : "") + "</span>";
            }
            return s;
          }
        },
        xAxis: { type: "category", data: ps, axisLine: { lineStyle: { color: "#ddd5c4" } } },
        yAxis: [
          { type: "value", name: sec.unit, splitLine: { lineStyle: { color: "#e7e0cf" } } },
          {
            type: "value", name: "同比%", position: "right",
            splitLine: { show: false },
            axisLabel: { formatter: "{value}%" }
          }
        ],
        series: series
      });
      chartInst.setOption(opt, true);
    }

    bq.addEventListener("click", function () {
      viewState[domId] = "quarter"; bq.className = "active"; by.className = "";
      renderFilterBar(); renderTable(); renderChart();
    });
    by.addEventListener("click", function () {
      viewState[domId] = "year"; by.className = "active"; bq.className = "";
      renderFilterBar(); renderTable(); renderChart();
    });

    renderFilterBar(); renderTable(); renderChart();
    window.addEventListener("resize", function () { if (chartInst) chartInst.resize(); });
  }

  /* ================= 区块：公司卡片 ================= */
  // 折叠文本块：超长约 120 字时默认收缩为一行摘要（点击展开/收起），保持卡片紧凑
  function foldText(labelHtml, text, isNote) {
    var LONG = 120;
    if (!text) return "";
    var open = isNote
      ? '<div class="co-note">' + labelHtml
      : '<div class="co-line"><b>' + labelHtml + "</b>";
    var close = "</div>";
    if (text.length <= LONG) return open + esc(text) + close;
    var summaryInner = (isNote ? labelHtml : "<b>" + labelHtml + "</b>") +
      esc(text.slice(0, 60)) + "… <span class='fold-more'>更多</span>";
    var bodyInner = (isNote ? labelHtml : "<b>" + labelHtml + "</b>") + esc(text);
    return '<details class="co-fold ' + (isNote ? "co-note" : "co-line") + '">' +
      "<summary>" + summaryInner + "</summary>" +
      '<div class="fold-body">' + bodyInner + "</div></details>";
  }

  function nextDisclosure(calendar, companyName) {
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var hits = calendar.filter(function (e) {
      return e.company === companyName && new Date(e.date + "T00:00:00") >= today;
    }).sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    return hits[0] || null;
  }

  function renderCompanyCards(container, commodity, section) {
    var sectionEl = el("div", "section");
    var head = el("div", "section-head");
    head.appendChild(el("div", "section-title", "公司卡片（" + esc(section.title) + "）"));
    sectionEl.appendChild(head);
    var grid = el("div", "cards-grid");
    section.companies.forEach(function (c) {
      var card = el("div", "co-card");
      var lp = latestPeriodWithData(c.data, section.quarters, section.years);
      var isAnnual = lp !== null && /^\d{4}$/.test(lp);
      var headHtml = '<div class="co-head"><span class="co-name">' + esc(c.name) +
        (c.est ? '<sup style="color:#b8912f">†</sup>' : "") +
        (c.project ? ' <span style="font-weight:400;color:#8b98a9;font-size:12px">· ' + esc(c.project) + "</span>" : "") +
        '</span><span class="co-country">' + esc(c.country || "") + "</span></div>";
      var mainHtml = '<div class="co-main">' +
        '<span class="co-value">' + (lp ? fmtNum(c.data[lp]) : "—") + "</span>" +
        '<span class="co-period">' + (lp ? esc(lp) + (isAnnual ? " · 年度披露" : "") + " · " + esc(section.unit) : "暂无数据") + "</span>" +
        '<span class="co-yoy">' + (lp && c.yoy[lp] !== undefined ? "同比 " + fmtPct(c.yoy[lp]) : "") + "</span></div>";
      var pendHtml = "";
      if (c.pending) {
        var pps = Object.keys(c.pending).sort(function (a, b) { return periodKey(a) - periodKey(b); });
        var lastP = pps[pps.length - 1];
        pendHtml = '<div class="co-line">⏳ <b>' + esc(lastP) + "</b>：" + esc(c.pending[lastP]) + "</div>";
      }
      // 事件标记（事故/停产等，来自 news.json affects 字段）
      var eventHtml = "";
      if (c.event_flag) {
        var inner = '⚠ <b>' + esc(c.event_flag.date || "") + "</b> " + esc(c.event_flag.note || "");
        eventHtml = '<div class="co-line event-flag">' +
          (c.event_flag.url ? '<a href="' + esc(c.event_flag.url) + '" target="_blank" rel="noopener">' + inner + "</a>" : inner) +
          "</div>";
      }
      // 成本行：纯数字/区间文本补单位，其余（含币种/口径说明）原样展示
      var costHtml = "";
      if (c.cost) {
        var costText = String(c.cost);
        if (/^[\d.,]+\s*[-–~]\s*[\d.,]+$/.test(costText) || /^[\d.,]+$/.test(costText)) costText += " $/t";
        costHtml = '<div class="co-line"><b>成本</b>：' + esc(costText) + "</div>";
      }
      var guideHtml = "";
      var g = c.guide2026;
      if (g) {
        var guideText = g.lo !== null
          ? (g.lo === g.hi ? fmtNum(g.lo) : fmtNum(g.lo) + "–" + fmtNum(g.hi)) + " " + g.unit
          : null;
        var barHtml = g.pct !== null
          ? '<div class="bar"><i style="width:' + Math.min(100, g.pct).toFixed(1) + '%"></i></div>'
          : "";
        var progHtml = "";
        if (g.completed !== null) {
          progHtml = '<div class="co-line">2026 已完成 <b>' + fmtNum(g.completed) + "</b>" +
            (g.annualized !== null ? " · 年化 " + fmtNum(g.annualized) : "") +
            (g.pct !== null ? " · <b>" + g.pct.toFixed(0) + "%</b> " + statusBadge(g.status) : "") +
            (g.status === "季节性，不年化" ? " · " + statusBadge(g.status) : "") +
            (g.note ? ' <span style="color:#5b6879">（' + esc(g.note) + "）</span>" : "") + "</div>";
        } else if (g.status) {
          progHtml = '<div class="co-line">' + statusBadge(g.status) + "</div>";
        }
        guideHtml = '<div class="co-guide"><div class="co-line"><b>2026 指引</b>：' + esc(g.raw) +
          (guideText ? "（≈" + esc(guideText) + "）" : "") + "</div>" + progHtml + barHtml + "</div>";
      }
      var reasonHtml = foldText("变化原因：", c.reason, false);
      // 估算/推算备注：默认收缩为一行摘要，点击展开全文
      var estHtml = "";
      if (c.est && c.est_note) {
        estHtml = '<details class="co-fold co-note"><summary>† 含估算/推算 · 点击展开</summary>' +
          '<div class="fold-body">† ' + esc(c.est_note) + "</div></details>";
      }
      var noteHtml = foldText("⚠ ", c.note, true);
      var nd = nextDisclosure(commodity.calendar, c.name);
      var nextHtml = nd
        ? '<div class="co-next">下次披露：' + (nd.approx ? "约 " : "") + esc(nd.date) + " · " + esc(nd.event) + "</div>"
        : "";
      card.innerHTML = headHtml + mainHtml + pendHtml + eventHtml + costHtml + guideHtml + reasonHtml + estHtml + noteHtml + nextHtml;
      grid.appendChild(card);
    });
    sectionEl.appendChild(grid);
    makeCollapsible(sectionEl, false);
    container.appendChild(sectionEl);
  }

  /* ================= 区块：2026 指引 vs 年化进度 ================= */
  var guideFilterState = {}; // commodity.key -> { comp:{rowKey:true=隐藏} }
  function renderGuideProgress(container, commodity) {
    var gp = commodity.guide_progress;
    if (!gp || !gp.rows || !gp.rows.length) return;
    var section = el("div", "section");
    var head = el("div", "section-head");
    head.appendChild(el("div", "section-title", "2026 指引 vs 年化进度"));
    head.appendChild(el("div", "section-sub",
      "仅列有指引或 2026 已披露数据的公司；仅 Q1 年化×4 / 含 Q2 年化×2，含季节性风险"));
    section.appendChild(head);
    var panel = el("div", "panel");
    var filterBar = el("div", "tbl-filter");
    panel.appendChild(filterBar);
    var wrap = el("div", "table-wrap");
    panel.appendChild(wrap);
    section.appendChild(panel);
    container.appendChild(section);

    var fs = guideFilterState[commodity.key] = guideFilterState[commodity.key] || { comp: {} };
    function rowKey(r) { return r.section + "|" + r.name; }
    // 一级公司归组：优先 sections 的 group 映射（Vedanta 系等），其次同名多行（镍/铜）
    var nameToGroup = {}, nameCount = {};
    commodity.sections.forEach(function (sec) {
      sec.companies.forEach(function (c) {
        if (c.group) nameToGroup[c.name] = c.group;
        nameCount[c.name] = (nameCount[c.name] || 0) + 1;
      });
    });
    function groupOf(r) {
      var base = r.name.split("·")[0];
      if (nameToGroup[base]) return nameToGroup[base];
      if ((nameCount[base] || 0) > 1) return base;
      return null;
    }
    var trueG = {};
    gp.rows.forEach(function (r) { var k = groupOf(r); if (k) trueG[k] = (trueG[k] || 0) + 1; });
    // 公司筛选分组：按一级公司归组，同组多行给整组开关
    var groups = [], byP = {};
    gp.rows.forEach(function (r) {
      var p = groupOf(r) || r.name.split("·")[0];
      if (!byP[p]) { byP[p] = { name: p, items: [] }; groups.push(byP[p]); }
      byP[p].items.push(r);
    });

    function renderTable() {
      var rows = gp.rows.filter(function (r) { return !fs.comp[rowKey(r)]; });
      // 渲染分组（保持原顺序）：集团组 = 父行小计 + 子行缩进；独立行带「独立」标
      var rgroups = [], rby = {};
      rows.forEach(function (r) {
        var k = groupOf(r) || r.name;
        if (!rby[k]) { rby[k] = { name: k, grp: groupOf(r), items: [] }; rgroups.push(rby[k]); }
        rby[k].items.push(r);
      });
      var html = '<table class="data-table"><thead><tr>' +
        "<th>板块</th><th>公司</th><th>FY2026 指引（原文）</th><th>指引折算</th>" +
        "<th>2026 已完成</th><th>年化</th><th>完成度</th><th>状态</th><th>2027 展望</th><th>备注</th></tr></thead><tbody>";
      rgroups.forEach(function (g) {
        var isGroup = g.grp && (trueG[g.grp] || 0) > 1;
        if (isGroup) {  // 集团小计父行：指引/已完成/年化求和，完成度按合计中枢重算
          var lo = 0, hi = 0, comp = 0, ann = 0, nlo = 0, ncomp = 0, nann = 0;
          g.items.forEach(function (r) {
            if (r.lo !== null) { lo += r.lo; nlo++; }
            if (r.hi !== null) hi += r.hi;
            if (r.completed !== null) { comp += r.completed; ncomp++; }
            if (r.annualized !== null) { ann += r.annualized; nann++; }
          });
          var mid = nlo ? (lo + hi) / 2 : null;
          var pct2 = (mid && nann) ? ann / mid * 100 : null;
          var st = pct2 === null ? null : (pct2 >= 100 ? "超出" : pct2 >= 90 ? "符合" : "不及");
          html += '<tr class="grp-row"><td>' + esc(g.items[0].section) + "</td><td><b>" + esc(g.name) + "</b>" +
            ' <span class="grp-n">集团小计 ' + g.items.length + " 项目</span></td>" +
            '<td class="flat">-</td><td>' +
            (nlo ? fmtNum(lo) + "–" + fmtNum(hi) + " " + esc(g.items[0].unit) + (nlo < g.items.length ? ' <small>部分</small>' : "") : '<span class="flat">-</span>') + "</td><td>" +
            (ncomp ? fmtNum(comp) : '<span class="flat">-</span>') + "</td><td>" +
            (nann ? fmtNum(ann) : '<span class="flat">-</span>') + "</td><td>" +
            (pct2 !== null ? "<b>" + pct2.toFixed(0) + "%</b>" : '<span class="flat">-</span>') + "</td><td>" +
            (st ? statusBadge(st) : '<span class="flat">-</span>') + "</td>" +
            '<td class="flat">-</td><td style="color:#5b6879">集团小计</td></tr>';
        }
        g.items.forEach(function (r) {
          var guideCvt = r.lo !== null
            ? (r.lo === r.hi ? fmtNum(r.lo) : fmtNum(r.lo) + "–" + fmtNum(r.hi)) + " " + esc(r.unit)
            : '<span class="flat">-</span>';
          var nameCell = esc(r.name) + (isGroup ? "" : ' <span class="solo-badge" title="独立公司：不属于任何集团分组">独立</span>');
          html += "<tr><td>" + esc(r.section) + "</td><td" + (isGroup ? ' class="child-cell"' : "") + ">" + nameCell + "</td>" +
            '<td style="text-align:left;font-family:inherit;white-space:normal;min-width:140px">' + esc(r.guide_raw || "-") + "</td>" +
            "<td>" + guideCvt + "</td>" +
            "<td>" + (r.completed !== null ? fmtNum(r.completed) : '<span class="flat">-</span>') + "</td>" +
            "<td>" + (r.annualized !== null ? fmtNum(r.annualized) : '<span class="flat">-</span>') + "</td>" +
            "<td>" + (r.pct !== null ? "<b>" + r.pct.toFixed(0) + "%</b>" : '<span class="flat">-</span>') + "</td>" +
            "<td>" + statusBadge(r.status) + "</td>" +
            '<td style="text-align:left;font-family:inherit;white-space:normal;min-width:150px;color:#7fa3d0">' + esc(r.fy2027 || "-") + "</td>" +
            '<td style="text-align:left;font-family:inherit;color:#5b6879">' + esc(r.note || "") + "</td></tr>";
        });
      });
      if (commodity.outlook2027 && commodity.outlook2027.total) {
        html += '<tr class="total-row outlook2027-row"><td colspan="2"><b>2027 品种总量展望</b>（' +
          esc(commodity.outlook2027.date) + "）</td>" +
          '<td colspan="8" style="text-align:left;font-family:inherit;white-space:normal">' +
          esc(commodity.outlook2027.total) + "</td></tr>";
      }
      wrap.innerHTML = html + "</tbody></table>";
    }

    function renderBar() {
      filterBar.innerHTML =
        '<details class="tf"><summary>公司筛选<b class="tf-n"></b></summary><div class="tf-body">' +
        groups.map(function (g, gi) {
          if (g.items.length === 1) {
            return '<label class="tf-item"><input type="checkbox" data-gcomp="' + esc(rowKey(g.items[0])) + '"> ' +
              esc(g.items[0].name) + "</label>";
          }
          return '<div class="tf-group"><label class="tf-master"><input type="checkbox" data-gmaster="' + gi + '"> ' +
            esc(g.name) + "</label>" +
            g.items.map(function (r) {
              return '<label class="tf-item"><input type="checkbox" data-gcomp="' + esc(rowKey(r)) + '"> ' +
                esc(r.name) + "</label>";
            }).join("") + "</div>";
        }).join("") +
        '</div></details>' +
        '<button class="tf-reset" type="button">重置筛选</button><span class="tf-hint"></span>';

      function refresh() {
        filterBar.querySelectorAll("[data-gcomp]").forEach(function (box) {
          box.checked = !fs.comp[box.getAttribute("data-gcomp")];
        });
        filterBar.querySelectorAll("[data-gmaster]").forEach(function (box) {
          var g = groups[+box.getAttribute("data-gmaster")];
          var off = g.items.filter(function (r) { return fs.comp[rowKey(r)]; }).length;
          box.checked = off === 0;
          box.indeterminate = off > 0 && off < g.items.length;
        });
        var nc = gp.rows.filter(function (r) { return fs.comp[rowKey(r)]; }).length;
        filterBar.querySelector(".tf-n").textContent = nc ? " 隐 " + nc : "";
        filterBar.querySelector(".tf-hint").textContent = nc
          ? "已隐藏 " + nc + " 行；2027 品种总量展望行始终保留"
          : "可按一级公司整组或单项目勾选隐藏";
      }
      filterBar.querySelectorAll("[data-gcomp]").forEach(function (box) {
        box.addEventListener("change", function () {
          var k = box.getAttribute("data-gcomp");
          if (box.checked) delete fs.comp[k]; else fs.comp[k] = true;
          refresh(); renderTable();
        });
      });
      filterBar.querySelectorAll("[data-gmaster]").forEach(function (box) {
        box.addEventListener("change", function () {
          var g = groups[+box.getAttribute("data-gmaster")];
          g.items.forEach(function (r) {
            var k = rowKey(r);
            if (box.checked) delete fs.comp[k]; else fs.comp[k] = true;
          });
          refresh(); renderTable();
        });
      });
      filterBar.querySelector(".tf-reset").addEventListener("click", function () {
        fs.comp = {}; refresh(); renderTable();
      });
      refresh();
    }

    renderBar(); renderTable();
  }

  /* ================= 区块：成本与资本开支 ================= */
  // 横向条形图（成本对比/成本曲线通用）：items 升序，labelFn 生成条端标签与 tooltip 主值
  function barChartPanel(cf, valueKey, labelFn) {
    var panel = el("div", "panel");
    panel.appendChild(el("div", "section-sub", cf.title));
    var div = el("div", "chart");
    div.style.height = Math.max(200, cf.items.length * 46 + 70) + "px";
    panel.appendChild(div);
    if (cf.footnotes && cf.footnotes.length) {
      var fn = el("div", "chart-footnotes");
      fn.innerHTML = cf.footnotes.map(function (t) { return "<div>· " + esc(t) + "</div>"; }).join("");
      panel.appendChild(fn);
    }
    if (!hasEcharts()) {
      // 图表库未加载（IE 兼容模式/老内核浏览器跑不动 echarts）：降级为纯 HTML 条形图
      var max = 0;
      cf.items.forEach(function (it) { if ((it[valueKey] || 0) > max) max = it[valueKey]; });
      var rowsHtml = cf.items.map(function (it) {
        var v = it[valueKey] || 0;
        var w = max ? Math.max(2, Math.round(v / max * 100)) : 0;
        return '<div class="hbar-row"><span class="hbar-name">' + esc(it.name) + "</span>" +
          '<span class="hbar-track"><span class="hbar-fill" style="width:' + w + '%"></span></span>' +
          '<span class="hbar-val">' + esc(labelFn(it)) + "</span></div>";
      }).join("");
      div.outerHTML = '<div class="hbar-fallback"><div class="hbar-note">简版条形图（当前浏览器内核不支持 ECharts，数据一致）</div>' + rowsHtml + "</div>";
      return panel;
    }
    // 延迟初始化：此时 panel 尚未挂进文档（clientWidth=0，echarts 会量成 0 宽图）。
    // requestAnimationFrame 回调在渲染完成后执行，容器已有真实尺寸。
    requestAnimationFrame(function () {
      if (!div.isConnected) return;
      var inst = echarts.init(div, null, { renderer: "svg" });
      inst.setOption(Object.assign(darkChartBase(), {
      grid: { left: 190, right: 110, top: 20, bottom: 34 },
      tooltip: {
        trigger: "item",
        backgroundColor: "#fdfaf3", borderColor: "#ddd5c4",
        textStyle: { color: "#2e2a22", fontSize: 12 },
        formatter: function (p) {
          var it = cf.items[p.dataIndex];
          return "<b>" + esc(it.name) + "</b><br/>" + esc(labelFn(it)) +
            (it.note ? "<br/><span style='color:#8b98a9'>" + esc(it.note) + "</span>" : "") +
            (it.raw ? "<br/><span style='color:#8b98a9'>原文：" + esc(it.raw) + "</span>" : "") +
            (it.est ? "<br/><span style='color:#b8912f'>估算/折算值</span>" : "");
        }
      },
      xAxis: { type: "value", name: cf.currency || "US$/t", splitLine: { lineStyle: { color: "#e7e0cf" } } },
      yAxis: {
        type: "category",
        data: cf.items.map(function (it) { return it.name; }),
        axisLine: { lineStyle: { color: "#ddd5c4" } },
        axisLabel: { fontSize: 12 }
      },
      series: [{
        type: "bar", barWidth: 18,
        data: cf.items.map(function (it, i) {
          return {
            value: it[valueKey],
            itemStyle: { color: it.est ? "#c9a52e" : PALETTE[i % PALETTE.length] }
          };
        }),
        label: {
          show: true, position: "right", color: "#2e2a22",
          fontFamily: "Consolas, monospace", fontSize: 12,
          formatter: function (p) { return labelFn(cf.items[p.dataIndex]); }
        }
      }]
      }));
      window.addEventListener("resize", function () { inst.resize(); });
    });
    return panel;
  }

  function renderCosts(container, commodity) {
    if (!commodity.capex && !commodity.costs && !commodity.cost_compare && !commodity.cost_curve) return;
    var section = el("div", "section");
    var head = el("div", "section-head");
    head.appendChild(el("div", "section-title", "成本与资本开支"));
    section.appendChild(head);

    // 成本对比图（锡）/ C1 成本曲线（锌）：横向条形，US$/t 升序
    if (commodity.cost_compare && commodity.cost_compare.items.length) {
      section.appendChild(barChartPanel(commodity.cost_compare, "value", function (it) { return it.label; }));
    }
    if (commodity.cost_curve && commodity.cost_curve.items.length) {
      var ccUnit = commodity.cost_curve.currency || "$/t";
      section.appendChild(barChartPanel(commodity.cost_curve, "mid", function (it) {
        return it.lo === it.hi
          ? fmtNum(it.lo) + " " + ccUnit
          : fmtNum(it.lo) + "–" + fmtNum(it.hi) + " " + ccUnit;
      }));
    }

    if (commodity.capex) {
      var panel = el("div", "panel");
      panel.appendChild(el("div", "section-sub", "资本开支（各公司币种/口径未折算，详见口径说明列）"));
      var wrap = el("div", "table-wrap");
      var html = '<table class="data-table"><thead><tr>';
      commodity.capex.headers.forEach(function (h) { html += "<th>" + esc(h || "") + "</th>"; });
      html += "</tr></thead><tbody>";
      commodity.capex.rows.forEach(function (row) {
        html += "<tr>";
        row.forEach(function (cell, i) {
          html += i < 2
            ? '<td style="text-align:left;font-family:inherit">' + esc(cell || "—") + "</td>"
            : '<td style="text-align:left;font-family:inherit;white-space:normal;min-width:160px">' + esc(cell || "—") + "</td>";
        });
        html += "</tr>";
      });
      wrap.innerHTML = html + "</tbody></table>";
      panel.appendChild(wrap);
      section.appendChild(panel);
    }

    if (commodity.costs) {
      var row = el("div", "chart-row");
      var charts = [];
      function lineChart(title, seriesDefs, currency) {
        var panel = el("div", "panel");
        panel.appendChild(el("div", "section-sub", title + "（" + currency + "）"));
        var div = el("div", "chart");
        div.style.height = "300px";
        panel.appendChild(div);
        row.appendChild(panel);
        if (!hasEcharts()) {
          div.outerHTML = '<div class="chart-fallback">图表加载失败（ECharts CDN 不可用）</div>';
          return;
        }
        // 同 barChartPanel：延迟到挂载后再初始化，避免 0 宽图
        requestAnimationFrame(function () {
          if (!div.isConnected) return;
          var inst = echarts.init(div, null, { renderer: "svg" });
          inst.setOption(Object.assign(darkChartBase(), {
          legend: { top: 4, textStyle: { color: "#6d6248", fontSize: 11 } },
          xAxis: { type: "category", data: seriesDefs[0].points.map(function (p) { return p.q; }), axisLabel: { rotate: 45 } },
          yAxis: { type: "value", name: currency, splitLine: { lineStyle: { color: "#e7e0cf" } } },
          series: seriesDefs.map(function (sd, i) {
            return {
              name: sd.name, type: "line", smooth: true, symbol: "circle", symbolSize: 5,
              itemStyle: { color: PALETTE[i] }, lineStyle: { width: 2 },
              data: sd.points.map(function (p) { return p.v; })
            };
          })
          }));
          charts.push(inst);
        });
      }
      var al = commodity.costs.alphamin;
      lineChart(al.title, [{
        name: "AISC",
        points: al.series.map(function (s) { return { q: s.q, v: s.aisc === undefined ? null : s.aisc }; })
      }], al.currency);
      var mx = commodity.costs.metalsx;
      lineChart(mx.title, [
        { name: "C1 现金成本", points: mx.series.map(function (s) { return { q: s.q, v: s.c1 === undefined ? null : s.c1 }; }) },
        { name: "AISC", points: mx.series.map(function (s) { return { q: s.q, v: s.aisc === undefined ? null : s.aisc }; }) }
      ], mx.currency);
      section.appendChild(row);
      window.addEventListener("resize", function () { charts.forEach(function (c) { c.resize(); }); });
    }
    container.appendChild(section);
  }

  /* ================= 区块：披露日历 ================= */
  function renderCalendarTable(entries, showCommodity) {
    var html = '<table class="cal-list"><thead><tr><th>日期</th>' +
      (showCommodity ? "<th>品种</th>" : "") +
      "<th>公司</th><th>事件</th><th>状态</th><th>倒计时</th></tr></thead><tbody>";
    var today = new Date(); today.setHours(0, 0, 0, 0);
    entries.forEach(function (e) {
      var d = new Date(e.date + "T00:00:00");
      var days = Math.round((d - today) / 86400000);
      var cd = days < 0
        ? '<span class="flat">已过 ' + (-days) + " 天</span>"
        : (days === 0 ? '<span class="days-left soon">今天</span>'
          : '<span class="days-left' + (days <= 14 ? " soon" : "") + '">还有 ' + days + " 天</span>");
      var badge = e.status === "待披露"
        ? '<span class="badge pending">待披露</span>'
        : '<span class="badge done">已披露待核</span>';
      html += "<tr><td class='date'>" + (e.approx ? "~" : "") + esc(e.date) + "</td>" +
        (showCommodity ? '<td><span class="badge commodity">' + esc(e.commodity) + "</span></td>" : "") +
        "<td>" + esc(e.company) + "</td><td>" + esc(e.event) + "</td><td>" + badge + "</td><td>" + cd + "</td></tr>";
    });
    return html + "</tbody></table>";
  }

  function renderChangelog(container, entries, limit) {
    var list = limit ? entries.slice(-limit).reverse() : entries.slice().reverse();
    var panel = el("div", "panel");
    list.forEach(function (c) {
      var item = el("div", "log-item");
      item.innerHTML = '<span class="log-date">' + esc(c.date || "") + "</span>" +
        '<div class="log-content">' + esc(c.content || "") + "</div>" +
        (c.source ? '<div class="log-source">来源：' + esc(c.source) + "</div>" : "");
      panel.appendChild(item);
    });
    container.appendChild(panel);
  }

  function renderCaliber(container, notes) {
    var det = el("details", "caliber");
    det.innerHTML = "<summary>口径说明（点击展开）</summary><ul>" +
      notes.map(function (n) { return "<li>" + esc(n) + "</li>"; }).join("") + "</ul>";
    container.appendChild(det);
  }

  /* ================= 区块：信息速递 ================= */
  var NEWS_BADGE_CLASS = { "锡": "c-tin", "锌": "c-zinc", "铝": "c-alu", "宏观": "c-macro" };
  var NEWS_IMPACT = {
    high: { cls: "impact-high", txt: "影响强" },
    mid: { cls: "impact-mid", txt: "影响中" },
    low: { cls: "impact-low", txt: "影响弱" },
  };
  function renderNewsSection(container, items, limit) {
    if (!items || !items.length) return;  // 无新闻不渲染该区块
    var shown = limit ? items.slice(0, limit) : items;
    var section = el("div", "section");
    var head = el("div", "section-head");
    head.appendChild(el("div", "section-title", "信息速递"));
    head.appendChild(el("div", "section-sub",
      limit && items.length > limit ? "最新 " + shown.length + " 条 / 共 " + items.length + " 条" : shown.length + " 条 · 最新在前"));
    section.appendChild(head);
    var grid = el("div", "news-grid");
    shown.forEach(function (n) {
      var card = el("div", "news-card");
      var bc = NEWS_BADGE_CLASS[n.commodity] || "c-macro";
      var imp = NEWS_IMPACT[n.impact] || NEWS_IMPACT.mid;
      var titleHtml = n.url
        ? '<a class="news-title" href="' + esc(n.url) + '" target="_blank" rel="noopener">' + esc(n.title) + "</a>"
        : '<span class="news-title">' + esc(n.title) + "</span>";
      var srcHtml = esc(n.source || "") +
        (n.url ? ' · <a href="' + esc(n.url) + '" target="_blank" rel="noopener">查看来源 →</a>' : "");
      card.innerHTML =
        '<div class="news-badge"><span class="' + bc + '">' +
        esc(n.commodity || "") + " · " + esc(n.category || "") + "</span>" +
        '<span class="news-impact ' + imp.cls + '">' + imp.txt + "</span></div>" +
        '<div class="news-body">' +
        '<div class="news-date">' + esc(n.date || "") + "</div>" +
        titleHtml +
        '<div class="news-summary">' + esc(n.summary || "") + "</div>" +
        '<div class="news-source">' + srcHtml + "</div></div>";
      grid.appendChild(card);
    });
    section.appendChild(grid);
    container.appendChild(section);
  }

  /* ================= 区块：近3个月供应扰动梳理 ================= */
  function renderDisruptions(container, commodity) {
    var items = commodity.disruptions || [];
    if (!items.length) return;  // 无数据不渲染该区块
    var section = el("div", "section");
    var head = el("div", "section-head");
    head.appendChild(el("div", "section-title", "近 3 个月供应扰动梳理"));
    head.appendChild(el("div", "section-sub",
      items.length + " 条 · 含影响仍在持续的更早事件（标「持续」）· 恢复进展与 2027 预期见末列"));
    section.appendChild(head);
    var panel = el("div", "panel");
    var wrap = el("div", "table-wrap");
    var html = '<table class="data-table dis-table"><thead><tr>' +
      "<th>时间</th><th>公司/资产</th><th>国家</th><th>类型</th><th>产能/规模</th><th>影响量级</th><th>恢复进展与 2027 预期</th><th>来源</th>" +
      "</tr></thead><tbody>";
    items.forEach(function (d) {
      var dirBadge = d.dir === "up"
        ? '<span class="dis-dir up">增</span>'
        : '<span class="dis-dir down">减</span>';
      var dateHtml = esc(d.date || "") + (d.ongoing ? ' <span class="dis-ongoing">持续</span>' : "");
      var srcHtml = esc(d.source || "") +
        (d.url ? ' <a href="' + esc(d.url) + '" target="_blank" rel="noopener">→</a>' : "");
      html += "<tr><td class='date'>" + dateHtml + "</td><td>" + esc(d.company || "") + "</td><td>" +
        esc(d.country || "") + "</td><td>" + dirBadge + " " + esc(d.type || "") + "</td><td>" +
        esc(d.capacity || "—") + "</td><td>" + esc(d.impact || "") + "</td><td>" +
        esc(d.recovery || "") + "</td><td>" + srcHtml + "</td></tr>";
    });
    wrap.innerHTML = html + "</tbody></table>";
    panel.appendChild(wrap);
    section.appendChild(panel);
    container.appendChild(section);
  }


  /* ================= 页面：总览 ================= */
  function renderOverview(main) {
    var DATA = window.SITE_DATA;
    // 品种卡片
    var sec1 = el("div", "section");
    var h1 = el("div", "section-head");
    h1.appendChild(el("div", "section-title", "品种总览"));
    sec1.appendChild(h1);
    var grid = el("div", "overview-grid");
    DATA.commodities.forEach(function (c) {
      var ov = c.overview || {};
      var card = el("div", "ov-card");
      var yoyHtml = ov.mine_yoy !== null && ov.mine_yoy !== undefined ? fmtPct(ov.mine_yoy) : '<span class="flat">—</span>';
      var pct = ov.total_companies ? (ov.disclosed / ov.total_companies * 100) : 0;
      var secDesc = (c.sections || []).map(function (s) { return s.title + " " + s.unit; }).join(" · ");
      card.innerHTML =
        "<h3>" + esc(c.name) + '<span class="ov-unit">' + esc(secDesc) + "</span></h3>" +
        '<div class="ov-row"><span class="k">最新披露季度</span><span class="v">' + esc(ov.latest_period || "—") + "</span></div>" +
        '<div class="ov-row"><span class="k">' + esc(ov.section_title || "矿山") + '产量合计（同口径同比）</span><span class="v">' + fmtNum(ov.mine_total) + " " + yoyHtml + "</span></div>" +
        '<div class="ov-row"><span class="k">最近更新</span><span class="v">' + esc(c.last_update || "—") + "</span></div>" +
        '<div class="ov-progress"><div class="bar"><i style="width:' + pct.toFixed(0) + '%"></i></div>' +
        '<div class="lbl"><span>' + esc(ov.latest_period || "") + " 披露进度</span><span>已披露 " + ov.disclosed + " / 共 " + ov.total_companies + " 家</span></div></div>";
      grid.appendChild(card);
    });
    sec1.appendChild(grid);
    main.appendChild(sec1);

    // 信息速递（全部新闻，最新在前；为空则不渲染）
    renderNewsSection(main, DATA.news || []);

    // 倒计时 + 最近日志
    var sec2 = el("div", "section");
    var twoCol = el("div", "two-col");
    var left = el("div");
    var hl = el("div", "section-head");
    hl.appendChild(el("div", "section-title", "财报披露倒计时"));
    left.appendChild(hl);
    var allCal = [];
    DATA.commodities.forEach(function (c) { allCal = allCal.concat(c.calendar); });
    var today = new Date(); today.setHours(0, 0, 0, 0);
    allCal.sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    var future = allCal.filter(function (e) { return new Date(e.date + "T00:00:00") >= today; });
    var recent = allCal.filter(function (e) { return new Date(e.date + "T00:00:00") < today; }).slice(-3).reverse();
    var panelL = el("div", "panel");
    var wrapL = el("div", "table-wrap");
    wrapL.innerHTML = renderCalendarTable(future.concat(recent), true);
    panelL.appendChild(wrapL);
    left.appendChild(panelL);
    twoCol.appendChild(left);

    var right = el("div");
    var hr = el("div", "section-head");
    hr.appendChild(el("div", "section-title", "最近更新日志"));
    right.appendChild(hr);
    var allLog = [];
    DATA.commodities.forEach(function (c) {
      c.changelog.forEach(function (l) { allLog.push({ date: l.date, content: l.content, source: l.source, commodity: c.name }); });
    });
    allLog.sort(function (a, b) { return (a.date || "") < (b.date || "") ? 1 : -1; });
    var panelR = el("div", "panel");
    allLog.slice(0, 5).forEach(function (l) {
      var item = el("div", "log-item");
      item.innerHTML = '<span class="log-date">' + esc(l.date || "") + '</span><span class="badge commodity">' + esc(l.commodity) + "</span>" +
        '<div class="log-content">' + esc((l.content || "").slice(0, 120)) + ((l.content || "").length > 120 ? "…" : "") + "</div>";
      panelR.appendChild(item);
    });
    right.appendChild(panelR);
    twoCol.appendChild(right);
    sec2.appendChild(twoCol);
    main.appendChild(sec2);
  }

  /* ================= 区块：品种综述（数据驱动 + 研究观点） ================= */
  function renderReview(container, commodity) {
    var rv = commodity.review;
    if (!rv || !rv.lines || !rv.lines.length) return;
    var section = el("div", "section");
    var head = el("div", "section-head");
    head.appendChild(el("div", "section-title", "品种综述"));
    var sub = "产量合计为板块内已披露公司直接求和（口径混杂，仅作量级）";
    if (rv.comment && rv.comment.date) sub += " · 观点更新于 " + rv.comment.date;
    head.appendChild(el("div", "section-sub", sub));
    section.appendChild(head);
    var panel = el("div", "panel");
    var html = "";
    rv.lines.forEach(function (l) {
      html += '<div class="co-line"><b>' + esc(l.section) + "</b>：" + esc(l.period) +
        " 合计 <b>" + fmtNum(l.total) + "</b> " + esc(l.unit) +
        (l.yoy !== null && l.yoy !== undefined ? "，同比 <b>" + fmtPct(l.yoy) + "</b>" : "") +
        "（披露 " + l.disclosed + "/" + l.n + " 家）</div>";
    });
    if (rv.n_guide) {
      var g = rv.guide;
      html += '<div class="co-line"><b>2026 指引进度</b>（' + rv.n_guide + " 家有数值指引）：" +
        '<span style="color:#3fb97c">超出 ' + g["超出"] + "</span> · " +
        '<span style="color:#b8912f">符合 ' + g["符合"] + "</span> · " +
        '<span style="color:#e05d5d">不及 ' + g["不及"] + "</span>（详见下方指引表）</div>";
    }
    // 品种级产量事件（事故/减产/复产，来自 news affects company="*"）
    if (commodity.commodity_events && commodity.commodity_events.length) {
      html += '<div class="co-line"><b>产量事件</b>：</div>' +
        commodity.commodity_events.map(function (ev) {
          var inner = "⚠ <b>" + esc(ev.date || "") + "</b> " + esc(ev.note || "");
          return '<div class="event-flag" style="margin:2px 0">' +
            (ev.url ? '<a href="' + esc(ev.url) + '" target="_blank" rel="noopener">' + inner + "</a>" : inner) + "</div>";
        }).join("");
    }
    if (rv.comment && rv.comment.lines && rv.comment.lines.length) {
      html += '<div class="review-comment"><div class="co-line"><b>研究观点</b>（' + esc(rv.comment.date) + "）：</div>" +
        rv.comment.lines.map(function (t) { return '<div class="review-line">' + esc(t) + "</div>"; }).join("") + "</div>";
    }
    if (rv.global && rv.global.text) {
      html += '<div class="review-comment"><div class="co-line"><b>跨品种供应端强弱对比</b>（' + esc(rv.global.date) + "）：</div>" +
        '<div class="review-line">' + esc(rv.global.text) + "</div></div>";
    }
    panel.innerHTML = html;
    section.appendChild(panel);
    container.appendChild(section);
  }

  /* ================= 页面：品种 ================= */
  function renderCommodity(main, commodity) {
    // 信息速递置顶（本品种 + 宏观；为空则不渲染）
    var newsItems = (window.SITE_DATA.news || []).filter(function (n) {
      return n.commodity === commodity.name || n.commodity === "宏观";
    });
    renderNewsSection(main, newsItems, 4);
    // 近3个月供应扰动梳理（减停产/复产事件表，含恢复与 2027 预期）
    renderDisruptions(main, commodity);
    // 品种综述紧随其后
    renderReview(main, commodity);
    // 数据板块按品种配置驱动（锡/锌=2 个板块，铝=3 个板块，可任意扩展）
    commodity.sections.forEach(function (sec) {
      renderProductionSection(main, sec, commodity.key, sec.key, commodity.default_view);
    });
    commodity.sections.forEach(function (sec) {
      renderCompanyCards(main, commodity, sec);
    });
    renderGuideProgress(main, commodity);
    try {
      renderCosts(main, commodity);
    } catch (err) {
      var errPanel = el("div", "panel");
      errPanel.style.color = "#ff9d9d";
      errPanel.textContent = "成本区块渲染失败：" + err.message + "（请截图发给维护者）";
      main.appendChild(errPanel);
    }

    // 披露日历
    var secC = el("div", "section");
    var hc = el("div", "section-head");
    hc.appendChild(el("div", "section-title", "披露日历"));
    secC.appendChild(hc);
    var panelC = el("div", "panel");
    var wrapC = el("div", "table-wrap");
    var cal = commodity.calendar.slice().sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    wrapC.innerHTML = renderCalendarTable(cal, false);
    panelC.appendChild(wrapC);
    secC.appendChild(panelC);
    makeCollapsible(secC, true);
    main.appendChild(secC);

    // 更新日志
    var secL = el("div", "section");
    var hl = el("div", "section-head");
    hl.appendChild(el("div", "section-title", "更新日志"));
    secL.appendChild(hl);
    renderChangelog(secL, commodity.changelog, 0);
    makeCollapsible(secL, true);
    main.appendChild(secL);

    // 口径说明
    var secN = el("div", "section");
    renderCaliber(secN, commodity.caliber_notes);
    main.appendChild(secN);
  }

  /* ================= 启动 ================= */
  function start() {
    var DATA = window.SITE_DATA;
    if (!DATA) {
      document.getElementById("main").innerHTML =
        '<div class="panel">数据文件 data/data.js 未找到，请先运行 build_site.py 生成数据。</div>';
      return;
    }
    document.getElementById("app").style.display = "";
    document.getElementById("build-time").textContent = "数据构建 " + DATA.build_time;
    document.getElementById("build-time-2").textContent = DATA.build_time +
      " · 图表引擎：" + (typeof echarts !== "undefined" ? "ECharts " + echarts.version : "未加载（降级模式）");

    var tabs = document.getElementById("tabs");
    var main = document.getElementById("main");
    var pages = [{ key: "overview", name: "总览" }].concat(
      DATA.commodities.map(function (c) { return { key: c.key, name: c.name, commodity: c }; })
    );
    function show(key) {
      Array.prototype.forEach.call(tabs.children, function (t) {
        t.className = "tab" + (t.dataset.key === key ? " active" : "");
      });
      main.innerHTML = "";
      if (key === "overview") renderOverview(main);
      else {
        var c = DATA.commodities.filter(function (x) { return x.key === key; })[0];
        renderCommodity(main, c);
      }
      window.scrollTo(0, 0);
    }
    pages.forEach(function (p) {
      var t = el("div", "tab", esc(p.name));
      t.dataset.key = p.key;
      t.addEventListener("click", function () { show(p.key); });
      tabs.appendChild(t);
    });
    show("overview");
  }

  document.addEventListener("DOMContentLoaded", function () {
    initGate(start);
  });
})();
