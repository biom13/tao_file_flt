/* taskpane.js - giao dien Excel add-in tao file .flt */
(function () {
  "use strict";

  var CHUNK_ROWS = 3000;
  var state = { result: null, files: null };

  function $(id) { return document.getElementById(id); }

  function setStatus(text, kind) {
    var el = $("status");
    el.textContent = text || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      setStatus("Add-in này chỉ chạy bên trong Excel. Hãy mở từ thẻ Home trong Excel.", "error");
      $("scan").disabled = true;
      return;
    }
    $("scan").addEventListener("click", onScan);
    $("saveZip").addEventListener("click", onSaveZip);
    if (typeof window.showDirectoryPicker === "function") {
      $("saveDir").hidden = false;
      $("saveDir").addEventListener("click", onSaveDir);
    }
  });

  // ---------- Doc workbook (chi sheet dang hien thi) ----------
  function readWorkbook() {
    return Excel.run(function (ctx) {
      var wss = ctx.workbook.worksheets;
      wss.load("items/name,items/visibility");
      return ctx.sync().then(function () {
        var visible = wss.items.filter(function (ws) { return ws.visibility === "Visible"; });
        var hiddenCount = wss.items.length - visible.length;
        var used = visible.map(function (ws) {
          var u = ws.getUsedRangeOrNullObject(true);
          u.load("isNullObject,rowIndex,rowCount,columnIndex,columnCount");
          return u;
        });
        return ctx.sync().then(function () {
          var out = [];
          var chain = Promise.resolve();
          visible.forEach(function (ws, i) {
            chain = chain.then(function () {
              var u = used[i];
              if (u.isNullObject) { out.push({ name: ws.name, rows: [] }); return; }
              var totalRows = u.rowIndex + u.rowCount;
              var cols = u.columnIndex + u.columnCount;
              var rows = [];
              var inner = Promise.resolve();
              for (var start = 0; start < totalRows; start += CHUNK_ROWS) {
                (function (s) {
                  inner = inner.then(function () {
                    var n = Math.min(CHUNK_ROWS, totalRows - s);
                    var r = ws.getRangeByIndexes(s, 0, n, cols);
                    r.load("values");
                    setStatus("Đang đọc sheet " + ws.name + " (" + Math.min(s + n, totalRows) + "/" + totalRows + " dòng)...");
                    return ctx.sync().then(function () {
                      for (var k = 0; k < r.values.length; k++) rows.push(r.values[k]);
                    });
                  });
                })(start);
              }
              return inner.then(function () { out.push({ name: ws.name, rows: rows }); });
            });
          });
          return chain.then(function () { return { sheets: out, hiddenCount: hiddenCount }; });
        });
      });
    });
  }

  // ---------- Quet ----------
  function onScan() {
    var btn = $("scan");
    btn.disabled = true;
    $("results").hidden = true;
    setStatus("Đang đọc workbook...");
    readWorkbook().then(function (data) {
      var result = FltCore.buildGroups(data.sheets);
      var files = [];
      var skippedLong = [];
      result.groups.forEach(function (g) {
        var built = FltCore.buildFltBytes(g.codes);
        built.skipped.forEach(function (c) { skippedLong.push(c); });
        files.push({ name: g.fileName, data: built.bytes, count: g.codes.length - built.skipped.length });
      });
      state.result = result;
      state.files = files;
      render(result, files, data.hiddenCount, skippedLong);
      btn.disabled = false;
    }).catch(function (err) {
      setStatus("Không đọc được workbook: " + (err && err.message ? err.message : err) +
        ". Kiểm tra workbook không ở chế độ chỉnh sửa ô rồi thử lại.", "error");
      btn.disabled = false;
    });
  }

  function render(result, files, hiddenCount, skippedLong) {
    var total = 0;
    files.forEach(function (f) { total += f.count; });

    if (files.length === 0) {
      setStatus("Không có dữ liệu hợp lệ. Kiểm tra tiêu đề cột nằm ở dòng 1 và sheet không bị ẩn.", "error");
      $("results").hidden = false;
      $("summary").textContent = "Không tạo được file nào";
      $("fileTable").hidden = true;
      $("saveZip").hidden = true;
      $("saveDir").hidden = true;
      renderSheets(result, hiddenCount);
      renderSingles([]);
      renderWarnings(result.warnings, skippedLong);
      return;
    }

    $("fileTable").hidden = false;
    $("saveZip").hidden = false;
    if (typeof window.showDirectoryPicker === "function") $("saveDir").hidden = false;

    setStatus("Đã quét xong. Chọn cách lưu bên dưới.", "ok");
    $("results").hidden = false;
    $("summary").textContent = files.length + " file .flt từ " + total + " mã nhân viên";

    var tbody = $("fileTable").querySelector("tbody");
    tbody.innerHTML = "";
    files.forEach(function (f) {
      var tr = document.createElement("tr");
      var a = document.createElement("td"); a.textContent = f.name;
      var b = document.createElement("td"); b.className = "num"; b.textContent = f.count;
      tr.appendChild(a); tr.appendChild(b);
      tbody.appendChild(tr);
    });

    renderSheets(result, hiddenCount);
    renderSingles(files);
    renderWarnings(result.warnings, skippedLong);
  }

  function renderSheets(result, hiddenCount) {
    var ul = $("sheetList");
    ul.innerHTML = "";
    result.sheets.forEach(function (s) {
      var li = document.createElement("li");
      var n = document.createElement("span"); n.textContent = s.name;
      var m = document.createElement("span"); m.className = "note";
      m.textContent = s.used ? s.note : "Bỏ qua: " + s.note;
      li.appendChild(n); li.appendChild(m);
      ul.appendChild(li);
    });
    if (hiddenCount > 0) {
      var li2 = document.createElement("li");
      var n2 = document.createElement("span"); n2.textContent = hiddenCount + " sheet ẩn";
      var m2 = document.createElement("span"); m2.className = "note"; m2.textContent = "Bỏ qua: sheet đang bị ẩn";
      li2.appendChild(n2); li2.appendChild(m2);
      ul.appendChild(li2);
    }
  }

  function renderSingles(files) {
    var ul = $("singleList");
    ul.innerHTML = "";
    files.forEach(function (f) {
      var li = document.createElement("li");
      var n = document.createElement("span"); n.textContent = f.name;
      var b = document.createElement("button"); b.type = "button"; b.textContent = "Tải";
      b.addEventListener("click", function () { downloadBytes(f.name, f.data); });
      li.appendChild(n); li.appendChild(b);
      ul.appendChild(li);
    });
  }

  function renderWarnings(warnings, skippedLong) {
    var box = $("warnings");
    box.innerHTML = "";
    var msgs = warnings.slice();
    if (skippedLong.length > 0) {
      msgs.push(skippedLong.length + " mã nhân viên dài hơn 255 byte nên không được ghi vào file .flt.");
    }
    box.hidden = msgs.length === 0;
    msgs.forEach(function (m) {
      var p = document.createElement("p"); p.textContent = m; box.appendChild(p);
    });
  }

  // ---------- Luu file ----------
  function downloadBytes(name, bytes) {
    var blob = new Blob([bytes], { type: "application/octet-stream" });
    if (navigator.msSaveBlob) { navigator.msSaveBlob(blob, name); return; }
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1500);
  }

  function onSaveZip() {
    if (!state.files || state.files.length === 0) return;
    var zip = FltCore.makeZip(state.files.map(function (f) { return { name: f.name, data: f.data }; }));
    var d = new Date();
    function p2(n) { return (n < 10 ? "0" : "") + n; }
    var stamp = d.getFullYear() + p2(d.getMonth() + 1) + p2(d.getDate()) + "_" + p2(d.getHours()) + p2(d.getMinutes());
    downloadBytes("FLT_result_" + stamp + ".zip", zip);
    setStatus("Đã gửi file .zip tới trình tải xuống. Nếu không thấy, dùng mục Tải từng file.", "ok");
  }

  function onSaveDir() {
    if (!state.files || state.files.length === 0) return;
    window.showDirectoryPicker({ mode: "readwrite" }).then(function (dir) {
      return dir.getDirectoryHandle("result", { create: true }).then(function (resDir) {
        var chain = Promise.resolve();
        state.files.forEach(function (f) {
          chain = chain.then(function () {
            return resDir.getFileHandle(f.name, { create: true }).then(function (fh) {
              return fh.createWritable().then(function (w) {
                return w.write(f.data).then(function () { return w.close(); });
              });
            });
          });
        });
        return chain;
      });
    }).then(function () {
      setStatus("Đã lưu " + state.files.length + " file vào thư mục result.", "ok");
    }).catch(function (err) {
      if (err && err.name === "AbortError") { setStatus("Đã hủy chọn thư mục."); return; }
      setStatus("Không lưu được vào thư mục: " + (err && err.message ? err.message : err) + ". Dùng nút Tải file .zip.", "error");
    });
  }
})();
