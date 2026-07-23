/* PDF 转 Word · 1:1 复原引擎
 * 原理：每页整页高清渲染为底图（图表、照片、线条像素级保留）；
 * 纯色背景上的文字从底图中擦除后，按原坐标用 Word 绝对定位文本框重建 ——
 * 页数、版式、图片与原文 1:1，且每个文字都可编辑。
 * 依赖：pdf.js 3.11.174（pdfjsLib）、docx 8.5.0（docx）。
 */
(function () {
  'use strict';

  var PDFJS_VERSION = '3.11.174';
  var PDFJS_DIST_ROOT = 'https://unpkg.com/pdfjs-dist@' + PDFJS_VERSION + '/';
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/' + PDFJS_VERSION + '/pdf.worker.min.js';

  var RUN_FONT = { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '宋体' };
  var SANS_FONT = { ascii: 'Arial', hAnsi: 'Arial', eastAsia: '微软雅黑' };

  // CMap 是 Type0/CID 中文 PDF 正确解码的关键。所有 PDF.js 路径必须统一使用同版本资源。
  function openPdf(arrayBuffer) {
    return pdfjsLib.getDocument({
      data: arrayBuffer,
      cMapUrl: PDFJS_DIST_ROOT + 'cmaps/',
      cMapPacked: true,
      standardFontDataUrl: PDFJS_DIST_ROOT + 'standard_fonts/',
      useSystemFonts: true,
      disableFontFace: false
    }).promise;
  }

  var dropzone = document.getElementById('dropzone');
  var fileInput = document.getElementById('fileInput');
  var fileListEl = document.getElementById('fileList');
  var actionsEl = document.getElementById('actions');
  var convertAllBtn = document.getElementById('convertAllBtn');
  var clearBtn = document.getElementById('clearBtn');

  // 任务队列：{ id, file, status, progress, message, blobUrl, outName, pageCount, warning }
  var tasks = [];
  var nextId = 1;
  var converting = false;

  // ---------- 文件选择与拖拽 ----------
  dropzone.addEventListener('click', function () { fileInput.click(); });
  fileInput.addEventListener('change', function () {
    addFiles(fileInput.files);
    fileInput.value = '';
  });
  ['dragenter', 'dragover'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });
  dropzone.addEventListener('drop', function (e) { addFiles(e.dataTransfer.files); });

  function addFiles(fileList) {
    var added = 0;
    Array.prototype.forEach.call(fileList, function (f) {
      var isPdf = /\.pdf$/i.test(f.name) || f.type === 'application/pdf';
      if (!isPdf) return;
      tasks.push({ id: nextId++, file: f, status: 'pending', progress: 0, message: '', blobUrl: null, outName: f.name.replace(/\.pdf$/i, '') + '.docx', warning: null });
      added++;
    });
    if (!added && fileList.length) alert('请选择 PDF 文件');
    render();
  }

  clearBtn.addEventListener('click', function () {
    if (converting) return;
    tasks.forEach(function (t) { if (t.blobUrl) URL.revokeObjectURL(t.blobUrl); });
    tasks = [];
    render();
  });

  convertAllBtn.addEventListener('click', function () {
    if (converting) return;
    runQueue();
  });

  // ---------- 渲染 ----------
  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    if (bytes < 1024 * 1024 * 1024) return (bytes / 1024 / 1024).toFixed(1) + ' MB';
    return (bytes / 1024 / 1024 / 1024).toFixed(2) + ' GB';
  }

  // 进度回调可能每秒触发几十次：渲染做节流，避免频繁重建整个列表 DOM
  var renderTimer = 0;
  function renderSoon() {
    if (renderTimer) return;
    renderTimer = setTimeout(function () { renderTimer = 0; render(); }, 100);
  }

  function render() {
    actionsEl.style.display = tasks.length ? 'flex' : 'none';
    convertAllBtn.disabled = converting || !tasks.some(function (t) { return t.status === 'pending' || t.status === 'err'; });
    convertAllBtn.textContent = converting ? '转换中…' : '开始转换';
    fileListEl.innerHTML = '';
    tasks.forEach(function (t) {
      var row = document.createElement('div');
      row.className = 'file-row';

      var main = document.createElement('div');
      main.style.flex = '1';
      main.style.minWidth = '0';

      var nameLine = document.createElement('div');
      nameLine.style.display = 'flex';
      nameLine.style.gap = '12px';
      var name = document.createElement('span');
      name.className = 'name';
      name.textContent = t.file.name;
      name.title = t.file.name;
      var size = document.createElement('span');
      size.className = 'size';
      size.textContent = formatSize(t.file.size);
      nameLine.appendChild(name);
      nameLine.appendChild(size);
      main.appendChild(nameLine);

      if (t.warning) {
        var warn = document.createElement('div');
        warn.className = 'warn';
        warn.textContent = '⚠ ' + t.warning;
        main.appendChild(warn);
      }
      if (t.status === 'busy') {
        var bar = document.createElement('div');
        bar.className = 'progress';
        var inner = document.createElement('div');
        inner.style.width = Math.round(t.progress * 100) + '%';
        bar.appendChild(inner);
        main.appendChild(bar);
      }
      row.appendChild(main);

      var status = document.createElement('span');
      status.className = 'status ' + (t.status === 'ok' ? 'ok' : t.status === 'err' ? 'err' : t.status === 'busy' ? 'busy' : '');
      status.textContent =
        t.status === 'pending' ? '待转换' :
        t.status === 'busy' ? (t.message || '解析中…') :
        t.status === 'ok' ? '完成（' + t.pageCount + ' 页）' :
        '失败：' + t.message;
      row.appendChild(status);

      if (t.status === 'ok') {
        var dl = document.createElement('button');
        dl.className = 'btn';
        dl.textContent = '下载 Word';
        dl.addEventListener('click', function () {
          var a = document.createElement('a');
          a.href = t.blobUrl;
          a.download = t.outName;
          document.body.appendChild(a);
          a.click();
          a.remove();
        });
        row.appendChild(dl);
      }

      var del = document.createElement('button');
      del.className = 'btn ghost';
      del.textContent = '移除';
      del.title = '从列表中移除该文件';
      del.disabled = converting;
      del.addEventListener('click', function () {
        if (converting) return;
        if (t.blobUrl) URL.revokeObjectURL(t.blobUrl);
        tasks = tasks.filter(function (x) { return x.id !== t.id; });
        render();
      });
      row.appendChild(del);

      fileListEl.appendChild(row);
    });
  }

  // ---------- 队列 ----------
  async function runQueue() {
    converting = true;
    render();
    try {
      for (var i = 0; i < tasks.length; i++) {
        var t = tasks[i];
        if (t.status !== 'pending' && t.status !== 'err') continue;
        await convertOne(t);
        render();
      }
    } finally {
      converting = false;
      render();
    }
  }

  async function convertOne(t) {
    t.status = 'busy';
    t.progress = 0;
    t.warning = null;
    t.message = '读取文件…';
    render();
    try {
      var buf = await t.file.arrayBuffer();
      var onProgress = function (done, total) {
        t.progress = done / total;
        renderSoon();
      };
      var onMessage = function (msg) {
        t.message = msg;
        renderSoon();
      };
      var result = await extractPdfOneToOne(buf, onProgress, onMessage);
      t.message = '生成 Word…';
      render();
      var blob = await buildDocxOneToOne(result.pages, { title: t.file.name });
      if (t.blobUrl) URL.revokeObjectURL(t.blobUrl);
      t.blobUrl = URL.createObjectURL(blob);
      t.status = 'ok';
      t.pageCount = result.pageCount;
      t.progress = 1;
      if (result.keptLines) {
        appendTaskWarning(t, result.keptLines + ' 个图片内/旋转/特殊字体的文字行保留在原位图像中（位置外观与原文一致）');
      }
      if (result.suspiciousPages) {
        appendTaskWarning(t, '检测到 ' + result.suspiciousPages + ' 页文字编码异常，已整页保留高清图像以避免乱码');
      }
    } catch (err) {
      console.error(err);
      t.status = 'err';
      t.message = friendlyError(err);
    }
  }

  function appendTaskWarning(task, message) {
    task.warning = task.warning ? task.warning + '；' + message : message;
  }

  function friendlyError(err) {
    var msg = (err && err.message) || String(err);
    if (/password/i.test(msg)) return 'PDF 已加密，无法读取';
    if (/Invalid PDF|corrupt|damaged/i.test(msg)) return 'PDF 文件损坏或格式异常';
    return '解析失败，请换其它浏览器重试';
  }

  // ---------- 工具 ----------
  // 只保留 XML 1.0 合法字符，同时修复常见 PDF 连字和汉字兼容码。孤立代理项替换为 U+FFFD，便于后续检测乱码页。
  function sanitize(value) {
    var s = String(value == null ? '' : value);
    var out = '';
    for (var i = 0; i < s.length;) {
      var cp = s.codePointAt(i);
      var units = cp > 0xFFFF ? 2 : 1;
      if (units === 1 && cp >= 0xD800 && cp <= 0xDFFF) {
        out += '�';
      } else if (cp === 0x09 || cp === 0x0A || cp === 0x0D ||
                 (cp >= 0x20 && cp <= 0xD7FF) ||
                 (cp >= 0xE000 && cp <= 0xFFFD) ||
                 (cp >= 0x10000 && cp <= 0x10FFFF)) {
        out += String.fromCodePoint(cp);
      }
      i += units;
    }
    out = out.replace(/\u00A0/g, ' ').replace(/\u00AD/g, '');
    out = out.replace(/[ﬀ-ﬆ]/g, function (ch) {
      return ({ 'ﬀ': 'ff', 'ﬁ': 'fi', 'ﬂ': 'fl', 'ﬃ': 'ffi',
        'ﬄ': 'ffl', 'ﬅ': 'st', 'ﬆ': 'st' })[ch] || ch;
    });
    // 部分中文 PDF 的 ToUnicode CMap 会返回康熙部首（如 U+2F42“⽂”）而不是普通汉字。
    // 两者外观近似但搜索、复制和 Word 校对不相等；只对该范围做兼容归一，避免改变全角字符等正常内容。
    out = out.replace(/[⼀-⿕]/g, function (ch) {
      return ch.normalize ? ch.normalize('NFKC') : ch;
    });
    return out.normalize ? out.normalize('NFC') : out;
  }

  function isCJK(code) {
    return (code >= 0x2E80 && code <= 0x9FFF) ||
      (code >= 0x3000 && code <= 0x303F) ||
      (code >= 0xF900 && code <= 0xFAFF) ||
      (code >= 0xFF00 && code <= 0xFFEF) ||
      (code >= 0x20000 && code <= 0x2FA1F) ||
      (code >= 0x3040 && code <= 0x30FF) ||
      (code >= 0xAC00 && code <= 0xD7AF);
  }

  var CLOSING_PUNCT = ',.;:!?%)]}、。，；：！？）】》”’';
  var OPENING_PUNCT = '([{（【《“‘';

  // 两个相邻文字碎片之间是否需要补空格：中文与中文之间不加，标点紧贴不加
  function needsSpace(prevStr, nextStr) {
    if (!prevStr || !nextStr) return false;
    if (/\s$/.test(prevStr) || /^\s/.test(nextStr)) return false;
    var prevChars = Array.from(prevStr);
    var nextChars = Array.from(nextStr);
    var ca = prevChars[prevChars.length - 1].codePointAt(0);
    var cb = nextChars[0].codePointAt(0);
    if (isCJK(ca) || isCJK(cb)) return false;
    if (CLOSING_PUNCT.indexOf(nextChars[0]) !== -1) return false;
    if (OPENING_PUNCT.indexOf(prevChars[prevChars.length - 1]) !== -1) return false;
    return true;
  }

  // 检测乱码：除私用区/替换符外，也覆盖常见 UTF-8 被 Latin-1/GBK 错误解码的特征。
  function textQualityOf(text) {
    var clean = sanitize(text);
    var bad = 0, total = 0;
    for (var ch of clean) {
      if (/\s/.test(ch)) continue;
      total++;
      var cp = ch.codePointAt(0);
      if (cp === 0xFFFD || (cp >= 0xE000 && cp <= 0xF8FF) || cp < 0x20) bad++;
    }
    var mojibake = clean.match(/(?:锟斤拷|烫烫烫|屯屯屯|\uFFFD|(?:Ã|Â)[\u0080-\u00BF]|\u00E2(?:\u0080|\u20AC)|\u00F0\u0178)/g);
    var mojibakeCount = mojibake ? mojibake.length : 0;
    var ratio = total ? (bad + mojibakeCount) / total : 0;
    return {
      text: clean,
      visible: total,
      ratio: ratio,
      suspicious: ratio > 0.008
    };
  }

  function garbleRatioOf(text) {
    return textQualityOf(text).ratio;
  }

  function normalizeTextItems(tc, viewport) {
    var U = pdfjsLib.Util;
    return tc.items.map(function (it) {
      var style = tc.styles[it.fontName] || {};
      var tx = U.transform(viewport.transform, it.transform);
      var size = Math.hypot(tx[2], tx[3]) || Math.hypot(tx[0], tx[1]) || 12;
      var family = String(style.fontFamily || '');
      var styleName = family + ' ' + String(it.fontName || '');
      var ascent = typeof style.ascent === 'number' && isFinite(style.ascent) && style.ascent > 0 && style.ascent < 4 ?
        style.ascent : 0.8;
      var descent = typeof style.descent === 'number' && isFinite(style.descent) && Math.abs(style.descent) < 4 ?
        Math.abs(style.descent) : 0.2;
      return {
        str: sanitize(it.str || ''),
        transform: tx,
        width: Math.abs((it.width || 0) * (viewport.scale || 1)),
        height: Math.abs((it.height || size) * (viewport.scale || 1)),
        size: size,
        fontName: it.fontName || '',
        fontFamily: family,
        bold: /bold|black|heavy|semibold|demi/i.test(styleName),
        italics: /italic|oblique/i.test(styleName),
        ascent: ascent,
        descent: descent,
        vertical: !!style.vertical,
        dir: it.dir || 'ltr',
        hasEOL: !!it.hasEOL,
        angle: Math.atan2(tx[1], tx[0])
      };
    });
  }

  // 把同一水平线上的文字碎片合并成行
  function groupLines(items) {
    var lines = [];
    items.forEach(function (it) {
      var str = it.str || '';
      if (!str) return;
      var x = it.transform[4];
      var y = it.transform[5];
      var size = it.size || Math.hypot(it.transform[2], it.transform[3]) || 12;
      var line = null;
      for (var i = 0; i < lines.length; i++) {
        // 容差取较大字号的 0.45 倍：上标（如 x² 的 ²）能并入正文行，而不会被拆成单独一行
        if (Math.abs(lines[i].y - y) <= Math.max(2, Math.max(size, lines[i].size) * 0.45)) { line = lines[i]; break; }
      }
      if (!line) { line = { y: y, size: size, parts: [] }; lines.push(line); }
      if (size > line.size) { line.size = size; line.y = y; }
      line.parts.push({
        x: x, y: y, str: str, width: it.width || 0, size: size,
        fontName: it.fontName || '', fontFamily: it.fontFamily || '', bold: !!it.bold, italics: !!it.italics,
        ascent: it.ascent == null ? 0.8 : it.ascent,
        descent: it.descent == null ? 0.2 : it.descent,
        vertical: !!it.vertical, dir: it.dir || 'ltr', angle: it.angle || 0,
        hasEOL: !!it.hasEOL
      });
    });
    lines.forEach(function (l) {
      l.parts.sort(function (a, b) { return a.x - b.x; });
      // PDF.js 常把源文件中的空格拆成独立 item。保留其位置用于恢复文字内容，
      // 但不把空白 item 自身做成 Word 文字框，也不让它扩大行的可见边界。
      var whitespaceParts = l.parts.filter(function (p) { return !p.str.trim(); });
      var contentParts = l.parts.filter(function (p) { return p.str.trim(); });
      var allParts = l.parts;
      l.parts = contentParts;
      if (!contentParts.length) {
        l.text = '';
        l.runs = [];
        return;
      }
      l.x = l.parts[0].x;
      l.xEnd = l.parts[0].x + l.parts[0].width;
      l.top = Infinity;
      l.bottom = -Infinity;
      var boldChars = 0, italicChars = 0, totalChars = 0;
      for (var i = 1; i < l.parts.length; i++) {
        var end = l.parts[i].x + l.parts[i].width;
        if (end > l.xEnd) l.xEnd = end;
      }
      l.parts.forEach(function (p) {
        l.top = Math.min(l.top, p.y - p.size * p.ascent);
        l.bottom = Math.max(l.bottom, p.y + p.size * p.descent);
        var chars = Array.from(p.str).length;
        totalChars += chars;
        if (p.bold) boldChars += chars;
        if (p.italics) italicChars += chars;
      });
      if (!isFinite(l.top)) l.top = l.y - l.size * 0.8;
      if (!isFinite(l.bottom)) l.bottom = l.y + l.size * 0.2;
      l.height = Math.max(l.size * 1.15, l.bottom - l.top);
      l.bold = totalChars ? boldChars / totalChars > 0.55 : false;
      l.italics = totalChars ? italicChars / totalChars > 0.55 : false;
      l.hasRotatedText = l.parts.some(function (p) {
        var a = Math.abs(p.angle) % Math.PI;
        return p.vertical || (a > 0.08 && Math.abs(a - Math.PI) > 0.08);
      });

      var text = '';
      var runs = [];
      for (var i = 0; i < l.parts.length; i++) {
        var p = l.parts[i];
        var prefix = '';
        if (i > 0) {
          var prev = l.parts[i - 1];
          var gap = p.x - (prev.x + prev.width);
          var gapLeft = prev.x + prev.width;
          var gapRight = p.x;
          var explicitSpace = whitespaceParts.some(function (space) {
            var center = space.x + Math.max(0, space.width) / 2;
            return center >= Math.min(gapLeft, gapRight) - 0.5 &&
              center <= Math.max(gapLeft, gapRight) + 0.5;
          });
          if (explicitSpace || (gap > l.size * 0.18 && needsSpace(prev.str, p.str))) prefix = ' ';
        }
        text += prefix + p.str;
        var runText = prefix + p.str;
        var last = runs[runs.length - 1];
        var closeEnough = last && p.x - last.xEnd <= Math.max(2, l.size * 0.35);
        var sameStyle = last && last.bold === p.bold && last.italics === p.italics &&
          Math.abs(last.size - p.size) < 0.2 && last.fontFamily === p.fontFamily &&
          last.fontName === p.fontName;
        if (closeEnough && sameStyle) {
          last.text += runText;
          last.xEnd = Math.max(last.xEnd, p.x + p.width);
          last.width = last.xEnd - last.x;
          last.hasEOL = last.hasEOL || p.hasEOL;
        } else {
          runs.push({
            // 新 run 会按 p.x 绝对定位，不能再把用于逻辑文本的补空格放进去，否则会二次右移。
            text: p.str, leadingSpace: prefix, x: p.x, xEnd: p.x + p.width, width: p.width,
            y: p.y, size: p.size, fontName: p.fontName, fontFamily: p.fontFamily,
            bold: p.bold, italics: p.italics, ascent: p.ascent, descent: p.descent,
            dir: p.dir, hasEOL: p.hasEOL
          });
        }
      }
      l.text = text.trim();
      l.runs = runs;
      l.hasEOL = allParts.some(function (p) { return p.hasEOL; });
    });
    return lines.filter(function (l) { return l.text; });
  }

  function sortByY(lines) {
    return lines.slice().sort(function (a, b) { return a.y - b.y; });
  }

  // 页面装配：先在文字碎片层面做多栏切分（XY-cut），再在各栏内部拼行。
  // 必须先切栏再拼行：同一水平线上左右两栏的文字 y 相同，先拼行会把两栏粘成一行。
  function assemblePage(items) {
    var lines = [];
    var regions = splitColumns(items, 0);
    if (regions.length === 1) regions = splitColumnsAroundSpanning(items) || regions;
    regions.forEach(function (regionItems, columnIndex) {
      var regionLines = sortByY(groupLines(regionItems));
      var left = Infinity, right = -Infinity;
      regionLines.forEach(function (l) { left = Math.min(left, l.x); right = Math.max(right, l.xEnd); });
      regionLines.forEach(function (l) {
        l.columnIndex = columnIndex;
        l.columnLeft = isFinite(left) ? left : l.x;
        l.columnRight = isFinite(right) ? right : l.xEnd;
        lines.push(l);
      });
    });
    return lines;
  }

  // 满宽标题/页眉/页脚会跨过栏间空白，不能让它们破坏双栏识别。
  function splitColumnsAroundSpanning(items) {
    if (items.length < 18) return null;
    var starts = items.map(function (it) { return it.transform[4]; });
    var ends = items.map(function (it) { return it.transform[4] + (it.width || 0); });
    var pageLeft = Math.min.apply(null, starts), pageRight = Math.max.apply(null, ends);
    var pageSpan = pageRight - pageLeft;
    if (pageSpan < 120) return null;

    var narrow = items.filter(function (it) { return (it.width || 0) < pageSpan * 0.58; });
    if (narrow.length < 12) return null;
    var sizes = narrow.map(function (it) { return it.size || 12; }).sort(function (a, b) { return a - b; });
    var minGap = Math.max(14, (sizes[Math.floor(sizes.length / 2)] || 12) * 1.4);
    var intervals = narrow.map(function (it) {
      return [it.transform[4], it.transform[4] + (it.width || 0)];
    }).sort(function (a, b) { return a[0] - b[0]; });
    var coveredEnd = -Infinity, best = null;
    intervals.forEach(function (iv) {
      if (coveredEnd > -Infinity && iv[0] - coveredEnd >= minGap &&
          (!best || iv[0] - coveredEnd > best[1] - best[0])) best = [coveredEnd, iv[0]];
      coveredEnd = Math.max(coveredEnd, iv[1]);
    });
    if (!best) return null;

    var splitX = (best[0] + best[1]) / 2;
    var left = [], right = [], spanning = [];
    items.forEach(function (it) {
      var x0 = it.transform[4], x1 = x0 + (it.width || 0);
      if ((x0 < splitX && x1 > splitX) || (it.width || 0) >= pageSpan * 0.58) spanning.push(it);
      else ((x0 + x1) / 2 < splitX ? left : right).push(it);
    });
    if (left.length < 6 || right.length < 6) return null;

    var bodyY = left.concat(right).map(function (it) { return it.transform[5]; });
    var minY = Math.min.apply(null, bodyY), maxY = Math.max.apply(null, bodyY);
    var midY = (minY + maxY) / 2;
    var top = [], bottom = [];
    spanning.forEach(function (it) { (it.transform[5] <= midY ? top : bottom).push(it); });
    return [top, left, right, bottom].filter(function (region) { return region.length; });
  }

  // 找到没有被任何碎片跨越的竖直空白带，把碎片切成左右两栏，递归处理
  function splitColumns(items, depth) {
    if (items.length < 12 || depth >= 2) return [items];

    var sizes = items.map(function (it) { return Math.hypot(it.transform[2], it.transform[3]) || 12; })
      .sort(function (a, b) { return a - b; });
    var mid = sizes[Math.floor(sizes.length / 2)] || 12;
    var minGap = Math.max(14, mid * 1.4);

    var intervals = items.map(function (it) {
      return [it.transform[4], it.transform[4] + (it.width || 0)];
    }).sort(function (a, b) { return a[0] - b[0]; });
    var coveredEnd = -Infinity, best = null;
    for (var i = 0; i < intervals.length; i++) {
      var s = intervals[i][0], e = intervals[i][1];
      if (coveredEnd > -Infinity && s - coveredEnd >= minGap) {
        if (!best || (s - coveredEnd) > (best[1] - best[0])) best = [coveredEnd, s];
      }
      if (e > coveredEnd) coveredEnd = e;
    }
    if (!best) return [items];

    var splitX = (best[0] + best[1]) / 2;
    var left = [], right = [];
    items.forEach(function (it) {
      var center = it.transform[4] + (it.width || 0) / 2;
      (center < splitX ? left : right).push(it);
    });
    if (left.length < 6 || right.length < 6) return [items];
    return splitColumns(left, depth + 1).concat(splitColumns(right, depth + 1));
  }

  var IMG_RENDER_SCALE = 2;
  function boundedRenderScale(page, requestedScale) {
    var base = page.getViewport({ scale: 1 });
    var scale = requestedScale > 0 && isFinite(requestedScale) ? requestedScale : 1;
    var maxPixels = 16000000;
    var maxDimension = 5200;
    scale = Math.min(scale, maxDimension / Math.max(1, base.width, base.height));
    scale = Math.min(scale, Math.sqrt(maxPixels / Math.max(1, base.width * base.height)));
    return scale;
  }

  function canvasToPngBlob(canvas) {
    return new Promise(function (resolve, reject) {
      canvas.toBlob(function (blob) {
        if (blob) resolve(blob);
        else reject(new Error('页面图像编码失败，可能是页面尺寸或内存超限'));
      }, 'image/png');
    });
  }

  function viewportRect(viewport, rect) {
    var pts = [
      viewport.convertToViewportPoint(rect[0], rect[1]),
      viewport.convertToViewportPoint(rect[0], rect[3]),
      viewport.convertToViewportPoint(rect[2], rect[1]),
      viewport.convertToViewportPoint(rect[2], rect[3])
    ];
    var xs = pts.map(function (pt) { return pt[0]; });
    var ys = pts.map(function (pt) { return pt[1]; });
    return [Math.min.apply(null, xs), Math.min.apply(null, ys), Math.max.apply(null, xs), Math.max.apply(null, ys)];
  }

  function rectIoU(a, b) {
    var x0 = Math.max(a[0], b[0]), y0 = Math.max(a[1], b[1]);
    var x1 = Math.min(a[2], b[2]), y1 = Math.min(a[3], b[3]);
    var intersection = Math.max(0, x1 - x0) * Math.max(0, y1 - y0);
    var areaA = Math.max(0, a[2] - a[0]) * Math.max(0, a[3] - a[1]);
    var areaB = Math.max(0, b[2] - b[0]) * Math.max(0, b[3] - b[1]);
    return intersection / Math.max(1, areaA + areaB - intersection);
  }

  async function renderPagePng(page, scale) {
    var baseViewport = page.getViewport({ scale: 1 });
    var actualScale = boundedRenderScale(page, scale);
    var viewport = page.getViewport({ scale: actualScale });
    var canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    try {
      var ctx = canvas.getContext('2d', { alpha: false });
      if (!ctx) throw new Error('无法创建页面渲染画布');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport: viewport, background: '#ffffff' }).promise;
      var blob = await canvasToPngBlob(canvas);
      return {
        data: new Uint8Array(await blob.arrayBuffer()),
        width: canvas.width,
        height: canvas.height,
        pageWidth: baseViewport.width,
        pageHeight: baseViewport.height
      };
    } finally {
      canvas.width = 0;
      canvas.height = 0;
    }
  }

  // 只取图片的显示区域（不渲染），用于判断文字是否压在图片上
  async function extractImageRects(page, baseViewport) {
    var opList = await page.getOperatorList();
    var OPS = pdfjsLib.OPS;
    var U = pdfjsLib.Util;
    var ctm = [1, 0, 0, 1, 0, 0];
    var stack = [];
    var rects = [];
    var hasInvisibleText = false;
    for (var i = 0; i < opList.fnArray.length; i++) {
      var fn = opList.fnArray[i], args = opList.argsArray[i];
      if (fn === OPS.save) {
        stack.push(ctm.slice());
      } else if (fn === OPS.restore) {
        if (stack.length) ctm = stack.pop();
      } else if (fn === OPS.paintFormXObjectBegin) {
        stack.push(ctm.slice());
        var formMatrix = args && Array.isArray(args[0]) ? args[0] : args;
        if (formMatrix && formMatrix.length === 6) ctm = U.transform(ctm, formMatrix);
      } else if (fn === OPS.paintFormXObjectEnd) {
        if (stack.length) ctm = stack.pop();
      } else if (fn === OPS.transform) {
        ctm = U.transform(ctm, args);
      } else if (fn === OPS.setTextRenderMode) {
        // Tr=3 表示不可见文字：典型特征是扫描页 + OCR 文字层盖在整页扫描图上
        if (args && args[0] === 3) hasInvisibleText = true;
      } else if (fn === OPS.paintImageXObject || fn === OPS.paintJpegXObject ||
                 fn === OPS.paintInlineImageXObject || fn === OPS.paintImageMaskXObject) {
        // 图片绘制在单位正方形上，四个角经 CTM 变换后取包围盒
        var pts = [[0, 0], [1, 0], [0, 1], [1, 1]].map(function (pt) { return U.applyTransform(pt, ctm); });
        var xs = pts.map(function (pt) { return pt[0]; });
        var ys = pts.map(function (pt) { return pt[1]; });
        var x0 = Math.min.apply(null, xs), x1 = Math.max.apply(null, xs);
        var y0 = Math.min.apply(null, ys), y1 = Math.max.apply(null, ys);
        if (x1 - x0 >= 8 && y1 - y0 >= 8) rects.push([x0, y0, x1, y1]);
      }
    }
    var normalized = [];
    rects.forEach(function (rect) {
      var vr = viewportRect(baseViewport, rect);
      if (!normalized.some(function (x) { return rectIoU(x, vr) > 0.96; })) {
        normalized.push(vr);
      }
    });
    return { rects: normalized, hasInvisibleText: hasInvisibleText };
  }

  // 文字行与图片区域重叠超过 30% 时，视为图内文字，保留在底图中
  function lineOverlapsRects(line, rects) {
    var x0 = line.x, y0 = line.top, x1 = line.xEnd, y1 = line.bottom;
    var lineArea = Math.max(1, (x1 - x0) * (y1 - y0));
    return rects.some(function (r) {
      var ix0 = Math.max(r[0], x0), iy0 = Math.max(r[1], y0);
      var ix1 = Math.min(r[2], x1), iy1 = Math.min(r[3], y1);
      return Math.max(0, ix1 - ix0) * Math.max(0, iy1 - iy0) / lineArea > 0.3;
    });
  }

  function colorHex(r, g, b) {
    return [r, g, b].map(function (value) {
      return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0');
    }).join('').toUpperCase();
  }

  function estimatedTextWidthPt(text, size) {
    return Array.from(text || '').reduce(function (sum, ch) {
      if (/\s/.test(ch)) return sum + size * 0.33;
      if (isCJK(ch.codePointAt(0))) return sum + size;
      if (/[ilI1.,:;'|!]/.test(ch)) return sum + size * 0.28;
      if (/[MW@#%&]/.test(ch)) return sum + size * 0.82;
      return sum + size * 0.54;
    }, 0);
  }

  function hybridClearRect(line, pageWidth, pageHeight) {
    var size = isFinite(line.size) && line.size > 0 ? line.size : 12;
    var baseline = isFinite(line.y) ? line.y : 0;
    var top = isFinite(line.top) ? line.top : baseline - size * 0.82;
    var bottom = isFinite(line.bottom) ? line.bottom : baseline + size * 0.22;
    var horizontalMargin = line.italics ? 2 : 1.2;
    var verticalMargin = 1.2;
    var sourceX = isFinite(line.x) ? line.x : 0;
    var sourceXEnd = isFinite(line.xEnd) && line.xEnd > sourceX ? line.xEnd : sourceX + Math.max(1, size * 0.5);
    var x0 = Math.max(0, sourceX - horizontalMargin);
    var y0 = Math.max(0, top - verticalMargin);
    var x1 = Math.min(pageWidth, sourceXEnd + horizontalMargin);
    var y1 = Math.min(pageHeight, bottom + verticalMargin);
    return { x: x0, y: y0, width: Math.max(1, x1 - x0), height: Math.max(1, y1 - y0) };
  }

  function hybridFrameRect(line, pageWidth, pageHeight) {
    var clearRect = hybridClearRect(line, pageWidth, pageHeight);
    var x = clearRect.x;
    var size = isFinite(line.size) && line.size > 0 ? line.size : 12;
    var sourceHeight = isFinite(line.height) && line.height > 0 ? line.height : clearRect.height;
    var sourceXEnd = isFinite(line.xEnd) && line.xEnd > x ? line.xEnd : x + clearRect.width;
    var padding = Math.max(18, size * 1.5);
    var right = line.hybridBlock ? sourceXEnd + padding : Math.max(
      sourceXEnd + padding,
      isFinite(line.columnRight) ? line.columnRight + padding : sourceXEnd + padding
    );
    right = Math.min(pageWidth, right);
    var height = Math.min(pageHeight - clearRect.y,
      Math.max(clearRect.height + 3, sourceHeight + 3, size * 1.65 + 3));
    return {
      x: x,
      y: clearRect.y,
      width: Math.max(1, right - x),
      height: Math.max(1, height)
    };
  }

  function rectToCanvasPixels(rect, scale, canvas) {
    var x0 = Math.max(0, Math.floor(rect.x * scale));
    var y0 = Math.max(0, Math.floor(rect.y * scale));
    var x1 = Math.min(canvas.width, Math.ceil((rect.x + rect.width) * scale));
    var y1 = Math.min(canvas.height, Math.ceil((rect.y + rect.height) * scale));
    return { x0: x0, y0: y0, x1: x1, y1: y1, width: x1 - x0, height: y1 - y0 };
  }

  function hybridLineUnsafeReason(line) {
    if (line.hasRotatedText) return 'rotated';
    var runs = line.runs && line.runs.length ? line.runs : [];
    var text = String(line.text || '');
    if (/[\uFFFD\uE000-\uF8FF\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(text)) {
      return 'garbled';
    }
    var unsafeFont = runs.some(function (run) {
      var name = String(run.fontFamily || '') + ' ' + String(run.fontName || '');
      return /symbol|dingbat|wingdings|webdings|math|cmr|cmmi|cmsy|cmex|stix|mt extra/i.test(name);
    });
    if (unsafeFont) return 'font';
    return '';
  }

  function hybridBlocksForLine(line) {
    line.hybridOverlappingRuns = false;
    var sourceRuns = line.runs && line.runs.length ? line.runs : [{
      text: line.text,
      leadingSpace: '',
      x: line.x,
      xEnd: line.xEnd,
      width: line.xEnd - line.x,
      y: line.y,
      size: line.size,
      fontName: '',
      fontFamily: '',
      bold: line.bold,
      italics: line.italics,
      ascent: 0.8,
      descent: 0.2,
      dir: 'ltr'
    }];
    var safeRuns = sourceRuns.map(function (run) {
      var text = String(run.text || '');
      if (!text.trim()) return null;
      var size = isFinite(run.size) && run.size > 0 ? run.size :
        (isFinite(line.size) && line.size > 0 ? line.size : 12);
      var y = isFinite(run.y) ? run.y : (isFinite(line.y) ? line.y : 0);
      var ascent = isFinite(run.ascent) ? run.ascent : 0.8;
      var descent = isFinite(run.descent) ? run.descent : 0.2;
      var x = isFinite(run.x) ? run.x : (isFinite(line.x) ? line.x : 0);
      var width = isFinite(run.width) && run.width > 0 ? run.width : Math.max(1, size * 0.5);
      var xEnd = isFinite(run.xEnd) && run.xEnd > x ? run.xEnd : x + width;
      return Object.assign({}, run, {
        text: text,
        leadingSpace: String(run.leadingSpace || ''),
        x: x,
        xEnd: xEnd,
        width: xEnd - x,
        y: y,
        size: size,
        ascent: ascent,
        descent: descent,
        fontName: run.fontName || '',
        fontFamily: run.fontFamily || '',
        bold: !!run.bold,
        italics: !!run.italics,
        dir: run.dir || 'ltr'
      });
    }).filter(Boolean).sort(function (a, b) { return a.x - b.x; });

    var blocks = [];
    safeRuns.forEach(function (run) {
      var block = blocks[blocks.length - 1];
      var previous = block && block.runs[block.runs.length - 1];
      var maxSize = previous ? Math.max(previous.size, run.size) : run.size;
      var gap = previous ? run.x - previous.xEnd : Infinity;
      var duplicateOverlay = previous && previous.text === run.text &&
        Math.abs(run.x - previous.x) <= Math.max(1, maxSize * 0.15) &&
        Math.abs(run.y - previous.y) <= Math.max(1, maxSize * 0.12);
      var severeOverlap = previous && gap < -maxSize * 0.55;
      if (duplicateOverlay || severeOverlap) line.hybridOverlappingRuns = true;
      var sameBaseline = previous && Math.abs(run.y - previous.y) <= Math.max(1.5, maxSize * 0.24);
      var closeEnough = previous && gap <= Math.max(2.5, maxSize * 0.42) && gap >= -maxSize * 0.35;
      var sameDirection = previous && previous.dir === run.dir;
      if (!block || !sameBaseline || !closeEnough || !sameDirection) {
        var firstRun = Object.assign({}, run, { leadingSpace: '' });
        blocks.push({
          text: run.text,
          x: run.x,
          xEnd: run.xEnd,
          y: run.y,
          top: run.y - run.size * run.ascent,
          bottom: run.y + run.size * run.descent,
          size: run.size,
          fontName: run.fontName,
          fontFamily: run.fontFamily,
          bold: !!run.bold,
          italics: !!run.italics,
          dir: run.dir,
          runs: [firstRun],
          columnLeft: line.columnLeft,
          columnRight: line.columnRight,
          hasRotatedText: line.hasRotatedText,
          hybridBlock: true
        });
        return;
      }
      block.runs.push(run);
      block.text += (run.leadingSpace || '') + run.text;
      block.xEnd = Math.max(block.xEnd, run.xEnd);
      block.top = Math.min(block.top, run.y - run.size * run.ascent);
      block.bottom = Math.max(block.bottom, run.y + run.size * run.descent);
      block.italics = block.italics || !!run.italics;
      if (run.size > block.size) {
        block.size = run.size;
        block.y = run.y;
      }
    });
    blocks.forEach(function (block) {
      block.height = Math.max(block.size * 1.15, block.bottom - block.top);
    });
    return blocks;
  }

  // 严格采样：外沿主色、四角和行/列连续像素共同判断背景是否纯净。
  function sampleLineAppearance(ctx, scale, line, pageWidth, pageHeight) {
    var clearRect = hybridClearRect(line, pageWidth, pageHeight);
    var pxRect = rectToCanvasPixels(clearRect, scale, ctx.canvas);
    var sw = pxRect.width, sh = pxRect.height;
    if (sw < 3 || sh < 3 || sw * sh > 1600000) return null;

    var data;
    try {
      data = ctx.getImageData(pxRect.x0, pxRect.y0, sw, sh).data;
    } catch (e) {
      return null;
    }

    var edgeSize = Math.max(1, Math.min(Math.floor(Math.min(sw, sh) / 4), Math.ceil(scale * 0.75)));
    var edgeBins = Object.create(null);
    var edgeTotal = 0;
    for (var y = 0; y < sh; y++) {
      for (var x = 0; x < sw; x++) {
        if (x >= edgeSize && x < sw - edgeSize && y >= edgeSize && y < sh - edgeSize) continue;
        var index = (y * sw + x) * 4;
        if (data[index + 3] < 128) continue;
        var edgeKey = ((data[index] >> 4) << 8) | ((data[index + 1] >> 4) << 4) | (data[index + 2] >> 4);
        var edgeBin = edgeBins[edgeKey] || (edgeBins[edgeKey] = { count: 0, r: 0, g: 0, b: 0 });
        edgeBin.count++;
        edgeBin.r += data[index];
        edgeBin.g += data[index + 1];
        edgeBin.b += data[index + 2];
        edgeTotal++;
      }
    }
    if (!edgeTotal) return null;

    var background = null;
    Object.keys(edgeBins).forEach(function (key) {
      if (!background || edgeBins[key].count > background.count) background = edgeBins[key];
    });
    if (!background || background.count / edgeTotal < 0.52) return null;
    var br = background.r / background.count;
    var bg = background.g / background.count;
    var bb = background.b / background.count;
    var bgLum = 0.2126 * br + 0.7152 * bg + 0.0722 * bb;

    var foregroundBins = Object.create(null);
    var rowForeground = new Array(sh).fill(0);
    var columnForeground = new Array(sw).fill(0);
    var columnRuns = new Array(sw).fill(0);
    var cornerSize = Math.max(1, Math.min(Math.floor(Math.min(sw, sh) / 3), Math.ceil(scale * 1.5)));
    var cornerTotals = [0, 0, 0, 0];
    var cornerBackground = [0, 0, 0, 0];
    var total = 0, uniform = 0, edgeUniform = 0, foregroundTotal = 0, brightSecondary = 0;
    var horizontalRun = 0, maxHorizontalRun = 0, maxVerticalRun = 0;
    for (var p = 0; p < data.length; p += 4) {
      if (data[p + 3] < 128) continue;
      var pixel = p / 4;
      var py = Math.floor(pixel / sw);
      var px = pixel - py * sw;
      if (px === 0) horizontalRun = 0;
      var dr = data[p] - br, dg = data[p + 1] - bg, db = data[p + 2] - bb;
      var distance2 = dr * dr + dg * dg + db * db;
      var isBackground = distance2 <= 42 * 42;
      var isEdge = px < edgeSize || px >= sw - edgeSize || py < edgeSize || py >= sh - edgeSize;
      if (isEdge && isBackground) edgeUniform++;
      var corner = -1;
      if (px < cornerSize && py < cornerSize) corner = 0;
      else if (px >= sw - cornerSize && py < cornerSize) corner = 1;
      else if (px < cornerSize && py >= sh - cornerSize) corner = 2;
      else if (px >= sw - cornerSize && py >= sh - cornerSize) corner = 3;
      if (corner >= 0) {
        cornerTotals[corner]++;
        if (isBackground) cornerBackground[corner]++;
      }
      total++;
      if (isBackground) {
        uniform++;
        horizontalRun = 0;
        columnRuns[px] = 0;
        continue;
      }
      foregroundTotal++;
      rowForeground[py]++;
      columnForeground[px]++;
      horizontalRun++;
      columnRuns[px]++;
      maxHorizontalRun = Math.max(maxHorizontalRun, horizontalRun);
      maxVerticalRun = Math.max(maxVerticalRun, columnRuns[px]);
      var lum = 0.2126 * data[p] + 0.7152 * data[p + 1] + 0.0722 * data[p + 2];
      if (bgLum > 200 && lum > 160) brightSecondary++;
      var fkey = ((data[p] >> 4) << 8) | ((data[p + 1] >> 4) << 4) | (data[p + 2] >> 4);
      var fbin = foregroundBins[fkey] || (foregroundBins[fkey] = {
        count: 0, r: 0, g: 0, b: 0, distance: 0
      });
      fbin.count++;
      fbin.r += data[p];
      fbin.g += data[p + 1];
      fbin.b += data[p + 2];
      fbin.distance += Math.sqrt(distance2);
    }
    if (!total || edgeUniform / edgeTotal < 0.74 || uniform / total < 0.50) return null;
    var cleanCorners = cornerTotals.filter(function (count, index) {
      return count && cornerBackground[index] / count >= 0.72;
    }).length;
    if (cleanCorners < 3 || foregroundTotal < Math.max(2, total * 0.002)) return null;
    if (foregroundTotal / total > 0.46 || brightSecondary / total > 0.08) return null;
    if (rowForeground.some(function (count) { return count / sw > 0.82; })) return null;
    if (columnForeground.some(function (count) { return count / sh > 0.86; })) return null;
    if (maxHorizontalRun > Math.max(36, (line.size || 12) * scale * 1.8)) return null;
    if (maxVerticalRun > Math.max(sh * 0.90, (line.size || 12) * scale * 2.1)) return null;

    var foreground = null, foregroundScore = -1;
    Object.keys(foregroundBins).forEach(function (key) {
      var candidate = foregroundBins[key];
      var score = candidate.count * (1 + Math.min(3, candidate.distance / candidate.count / 64));
      if (score > foregroundScore) {
        foreground = candidate;
        foregroundScore = score;
      }
    });
    if (!foreground) return null;
    return {
      fill: colorHex(br, bg, bb),
      color: colorHex(foreground.r / foreground.count, foreground.g / foreground.count,
        foreground.b / foreground.count),
      clearRect: clearRect,
      frameRect: hybridFrameRect(line, pageWidth, pageHeight)
    };
  }

  // 宽松采样：严格采样失败时启用。取边缘主色作背景、最深色作文字色，不设一致性门槛，
  // 保证任何文字行都能从底图擦除并重建为可编辑文本框。
  function sampleLineAppearanceRelaxed(ctx, scale, line, pageWidth, pageHeight) {
    var clearRect = hybridClearRect(line, pageWidth, pageHeight);
    var pxRect = rectToCanvasPixels(clearRect, scale, ctx.canvas);
    var sw = pxRect.width, sh = pxRect.height;
    if (sw < 3 || sh < 3) return null;

    var data;
    try {
      data = ctx.getImageData(pxRect.x0, pxRect.y0, sw, sh).data;
    } catch (e) {
      return null;
    }

    var edgeSize = Math.max(1, Math.floor(Math.min(sw, sh) / 4));
    var edgeBins = Object.create(null);
    var edgeTotal = 0;
    var fg = null, fgLum = Infinity;
    for (var y = 0; y < sh; y++) {
      for (var x = 0; x < sw; x++) {
        var index = (y * sw + x) * 4;
        if (data[index + 3] < 128) continue;
        var isEdge = x < edgeSize || x >= sw - edgeSize || y < edgeSize || y >= sh - edgeSize;
        if (isEdge) {
          var key = ((data[index] >> 4) << 8) | ((data[index + 1] >> 4) << 4) | (data[index + 2] >> 4);
          var bin = edgeBins[key] || (edgeBins[key] = { count: 0, r: 0, g: 0, b: 0 });
          bin.count++;
          bin.r += data[index];
          bin.g += data[index + 1];
          bin.b += data[index + 2];
          edgeTotal++;
        }
        var lum = 0.2126 * data[index] + 0.7152 * data[index + 1] + 0.0722 * data[index + 2];
        if (lum < fgLum) {
          fgLum = lum;
          fg = { r: data[index], g: data[index + 1], b: data[index + 2] };
        }
      }
    }
    if (!edgeTotal) return null;
    var background = null;
    Object.keys(edgeBins).forEach(function (key) {
      if (!background || edgeBins[key].count > background.count) background = edgeBins[key];
    });
    return {
      fill: colorHex(background.r / background.count, background.g / background.count,
        background.b / background.count),
      color: fg ? colorHex(fg.r, fg.g, fg.b) : '000000',
      clearRect: clearRect,
      frameRect: hybridFrameRect(line, pageWidth, pageHeight)
    };
  }

  async function renderHybridPage(page, lines, baseViewport, scale, imgRects) {
    var actualScale = boundedRenderScale(page, scale);
    var viewport = page.getViewport({ scale: actualScale });
    var canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    try {
      var ctx = canvas.getContext('2d', { alpha: false, willReadFrequently: true });
      if (!ctx) throw new Error('无法创建页面渲染画布');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport: viewport, background: '#ffffff' }).promise;

      // 覆盖全页的图片是页面背景（有些生成器会嵌一张全页空白底图），
      // 不能让它把整页文字都误判为“图内文字”；只有局部内容图才需要把文字留在图中
      var pageArea = Math.max(1, baseViewport.width * baseViewport.height);
      var contentRects = (imgRects || []).filter(function (r) {
        return (r[2] - r[0]) * (r[3] - r[1]) / pageArea < 0.8;
      });

      lines.forEach(function (line) {
        line.hybridSourceBlocks = hybridBlocksForLine(line);
        line.hybridRejectReason = lineOverlapsRects(line, contentRects) ? 'image' : hybridLineUnsafeReason(line);
        if (line.hybridOverlappingRuns) line.hybridRejectReason = 'overlap';
        line.hybridBlocks = [];
        if (line.hybridRejectReason) {
          line.hybridEditable = false;
          return;
        }
        line.hybridSourceBlocks.forEach(function (block) {
          var appearance = sampleLineAppearance(ctx, actualScale, block, baseViewport.width, baseViewport.height) ||
            sampleLineAppearanceRelaxed(ctx, actualScale, block, baseViewport.width, baseViewport.height);
          if (!appearance) return;
          block.hybridEditable = true;
          block.backgroundFill = appearance.fill;
          block.textColor = appearance.color;
          block.hybridClearRect = appearance.clearRect;
          block.hybridRect = appearance.frameRect;
          line.hybridBlocks.push(block);
        });
        line.hybridEditable = line.hybridBlocks.length > 0;
      });

      // 必须在全部行完成采样后再清除原字，防止相邻行的采样读到已经修改过的像素。
      ctx.save();
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      lines.forEach(function (line) {
        (line.hybridBlocks || []).forEach(function (block) {
          var rect = rectToCanvasPixels(block.hybridClearRect, actualScale, canvas);
          ctx.fillStyle = '#' + block.backgroundFill;
          ctx.fillRect(rect.x0, rect.y0, rect.width, rect.height);
        });
      });
      ctx.restore();

      var blob = await canvasToPngBlob(canvas);
      return {
        data: new Uint8Array(await blob.arrayBuffer()),
        width: canvas.width,
        height: canvas.height,
        pageWidth: baseViewport.width,
        pageHeight: baseViewport.height
      };
    } finally {
      canvas.width = 0;
      canvas.height = 0;
    }
  }

  // 1:1 提取：每页渲染高清底图，文字逐块擦除并标记为可编辑重建；
  // 压在图片上的文字（图内标签）保留在底图中；文字编码异常的页整页保真。
  async function extractPdfOneToOne(arrayBuffer, onProgress, onMessage) {
    var pdf = await openPdf(arrayBuffer);
    var pages = [];
    var lineCount = 0;
    var editableBlocks = 0;
    var keptLines = 0;
    var suspiciousPages = 0;
    try {
      for (var p = 1; p <= pdf.numPages; p++) {
        onMessage('精准重建 ' + p + '/' + pdf.numPages + ' 页');
        var page = await pdf.getPage(p);
        var viewport = page.getViewport({ scale: 1 });
        var tc = await page.getTextContent();
        var lines = assemblePage(normalizeTextItems(tc, viewport));
        var pageText = lines.map(function (line) { return line.text; }).join('\n');
        var quality = textQualityOf(pageText);
        var raster;
        if (quality.suspicious && quality.visible) {
          suspiciousPages++;
          lines.forEach(function (line) {
            line.hybridEditable = false;
            line.hybridBlocks = [];
          });
          raster = await renderPagePng(page, IMG_RENDER_SCALE);
          keptLines += lines.length;
        } else {
          var imgInfo = await extractImageRects(page, viewport);
          if (imgInfo.hasInvisibleText) {
            // 扫描页 + 不可见 OCR 文字层：文字层位置常与扫描图有偏差且无法清除底层图像，
            // 整页保真为高清图像，避免重影和错位
            suspiciousPages++;
            lines.forEach(function (line) {
              line.hybridEditable = false;
              line.hybridBlocks = [];
            });
            raster = await renderPagePng(page, IMG_RENDER_SCALE);
          } else {
            raster = await renderHybridPage(page, lines, viewport, IMG_RENDER_SCALE, imgInfo.rects);
          }
        }
        lineCount += lines.length;
        editableBlocks += lines.reduce(function (sum, line) {
          return sum + (line.hybridBlocks || []).length;
        }, 0);
        keptLines += lines.filter(function (line) { return !line.hybridEditable; }).length;
        pages.push({
          width: viewport.width,
          height: viewport.height,
          lines: lines,
          raster: raster,
          pageIndex: p - 1
        });
        onProgress(p, pdf.numPages);
      }
      return {
        pages: pages,
        pageCount: pdf.numPages,
        lineCount: lineCount,
        editableBlocks: editableBlocks,
        keptLines: keptLines,
        suspiciousPages: suspiciousPages
      };
    } finally {
      pdf.destroy(); // 释放 pdf.js 内部缓存，批量转换大文件时避免内存堆积
    }
  }

  function pageProperties(widthPt, heightPt, marginPt) {
    var landscape = widthPt > heightPt;
    var shortPt = Math.min(widthPt, heightPt);
    var longPt = Math.max(widthPt, heightPt);
    var m = Math.max(0, marginPt || 0) * 20;
    return {
      type: docx.SectionType.NEXT_PAGE,
      page: {
        size: {
          width: Math.round(shortPt * 20),
          height: Math.round(longPt * 20),
          orientation: landscape ? docx.PageOrientation.LANDSCAPE : docx.PageOrientation.PORTRAIT
        },
        margin: { top: m, bottom: m, left: m, right: m, header: 0, footer: 0, gutter: 0 }
      }
    };
  }

  // ---------- 生成 Word ----------
  function documentStyles() {
    return {
      default: {
        document: {
          run: {
            font: RUN_FONT,
            size: 24,
            language: { value: 'en-US', eastAsia: 'zh-CN' }
          },
          paragraph: {
            spacing: { after: 120, line: 276 },
            widowControl: true
          }
        }
      }
    };
  }

  function fontForPdfRun(run) {
    var family = String(run.fontFamily || '') + ' ' + String(run.fontName || '');
    return /sans|gothic|hei|yahei|arial|helvetica/i.test(family) ? SANS_FONT : RUN_FONT;
  }

  function floatingImageRun(data, xPt, yPt, widthPt, heightPt, behindDocument, zIndex) {
    return new docx.ImageRun({
      type: 'png',
      data: data,
      transformation: {
        // Word 页面尺寸只精确到 1/20 磅；图片 extent 同步量化到同一 twip，避免底图边缘出现极细白缝。
        width: Math.max(1, Math.round(widthPt * 20) / 15),
        height: Math.max(1, Math.round(heightPt * 20) / 15)
      },
      floating: {
        horizontalPosition: {
          relative: docx.HorizontalPositionRelativeFrom.PAGE,
          offset: Math.max(0, Math.round(xPt * 12700))
        },
        verticalPosition: {
          relative: docx.VerticalPositionRelativeFrom.PAGE,
          offset: Math.max(0, Math.round(yPt * 12700))
        },
        allowOverlap: true,
        behindDocument: !!behindDocument,
        layoutInCell: true,
        lockAnchor: true,
        margins: { top: 0, right: 0, bottom: 0, left: 0 },
        wrap: { type: docx.TextWrappingType.NONE },
        zIndex: zIndex || 1
      }
    });
  }

  var wordMeasureContext = null;
  function measuredWordTextWidthPt(text, size, run) {
    try {
      if (!wordMeasureContext) {
        var canvas = document.createElement('canvas');
        canvas.width = 1;
        canvas.height = 1;
        wordMeasureContext = canvas.getContext('2d');
      }
      if (!wordMeasureContext) return estimatedTextWidthPt(text, size);
      var font = fontForPdfRun(run || {});
      var ascii = font.ascii || font.hAnsi || 'Times New Roman';
      var eastAsia = font.eastAsia || '宋体';
      wordMeasureContext.font = (run && run.italics ? 'italic ' : '') +
        (run && run.bold ? '700 ' : '400 ') + Math.max(1, size) + 'px "' + ascii + '", "' + eastAsia + '"';
      var measured = wordMeasureContext.measureText(String(text || '')).width;
      return measured > 0 && isFinite(measured) ? measured : estimatedTextWidthPt(text, size);
    } catch (e) {
      return estimatedTextWidthPt(text, size);
    }
  }

  function hybridTextScalePercent(line, sourceRuns) {
    var actualWidth = Math.max(1, line.xEnd - line.x);
    var measuredWidth = sourceRuns.reduce(function (sum, run) {
      var text = (run.leadingSpace || '') + (run.text || '');
      return sum + measuredWordTextWidthPt(text, run.size || line.size || 12, run);
    }, 0);
    if (!(measuredWidth > 0)) return 100;
    return Math.max(50, Math.min(200, Math.round(actualWidth / measuredWidth * 100)));
  }

  function hybridLineTextRuns(line) {
    var sourceRuns = line.runs && line.runs.length ? line.runs : [{
      text: line.text,
      leadingSpace: '',
      size: line.size,
      fontName: '',
      fontFamily: '',
      bold: line.bold,
      italics: line.italics,
      y: line.y,
      dir: 'ltr'
    }];
    var specs = sourceRuns.map(function (run) {
      var deltaPt = (line.y || 0) - (run.y == null ? line.y || 0 : run.y);
      return {
        text: (run.leadingSpace || '') + (run.text || ''),
        run: run,
        position: Math.abs(deltaPt) >= 0.6 ? Math.max(-30, Math.min(30, Math.round(deltaPt * 2))) : null
      };
    }).filter(function (spec) { return spec.text; });
    if (specs.map(function (spec) { return spec.text; }).join('').trim() !== (line.text || '').trim()) {
      specs = [{
        text: line.text || '',
        run: sourceRuns[0],
        position: null
      }];
    }
    var scalePercent = hybridTextScalePercent(line, sourceRuns);
    return specs.map(function (spec) {
      var options = {
        text: sanitize(spec.text),
        size: Math.max(8, Math.min(120, Math.round((spec.run.size || line.size || 12) * 2))),
        bold: !!spec.run.bold,
        italics: !!spec.run.italics,
        rightToLeft: spec.run.dir === 'rtl',
        font: fontForPdfRun(spec.run),
        color: line.textColor || '000000',
        scale: scalePercent,
        language: { value: 'en-US', eastAsia: 'zh-CN' }
      };
      if (spec.position !== null) options.position = spec.position;
      return new docx.TextRun(options);
    });
  }

  function hybridTextParagraph(line, pg) {
    var rect = line.hybridRect || hybridFrameRect(line, pg.width, pg.height);
    var x = isFinite(rect.x) ? Math.max(0, rect.x) : 0;
    var y = isFinite(rect.y) ? Math.max(0, rect.y) : 0;
    var width = isFinite(rect.width) && rect.width > 0 ? rect.width : Math.max(18, line.size || 12);
    var height = isFinite(rect.height) && rect.height > 0 ? rect.height : Math.max(18, (line.size || 12) * 1.65 + 3);
    return new docx.Paragraph({
      frame: {
        type: 'absolute',
        position: { x: Math.round(x * 20), y: Math.round(y * 20) },
        width: Math.max(20, Math.round(width * 20)),
        height: Math.max(20, Math.round(height * 20)),
        anchor: {
          horizontal: docx.FrameAnchorType.PAGE,
          vertical: docx.FrameAnchorType.PAGE
        },
        space: { horizontal: 0, vertical: 0 },
        anchorLock: true,
        wrap: docx.FrameWrap.NONE,
        rule: docx.HeightRule.EXACT
      },
      spacing: { before: 0, after: 0 },
      widowControl: false,
      overflowPunctuation: true,
      autoSpaceEastAsianText: false,
      children: hybridLineTextRuns(line)
    });
  }

  // 底图已在 Canvas 中清除了可编辑行的原字；Word 仅负责叠加文字，避免重影。
  // 每页一节、页边距为 0、元素全部绝对定位 —— Word 页数与 PDF 页数严格一致。
  function buildDocxOneToOne(pages, opts) {
    var sections = pages.map(function (pg) {
      var children = [new docx.Paragraph({
        spacing: { before: 0, after: 0 },
        children: [floatingImageRun(pg.raster.data, 0, 0, pg.width, pg.height, true, 1)]
      })];
      (pg.lines || []).forEach(function (line) {
        (line.hybridBlocks || []).forEach(function (block) {
          children.push(hybridTextParagraph(block, pg));
        });
      });
      children.push(new docx.Paragraph({
        spacing: { before: 0, after: 0, line: 20, lineRule: docx.LineRuleType.EXACT },
        children: [new docx.TextRun({ text: '', size: 2 })]
      }));
      return {
        properties: pageProperties(pg.width, pg.height, 0),
        children: children
      };
    });
    var doc = new docx.Document({
      creator: 'PDF转Word',
      title: opts.title || 'converted',
      styles: documentStyles(),
      sections: sections
    });
    return docx.Packer.toBlob(doc);
  }
})();
