/* knowledgevba — SJIS copy hook for VBA module pages
 *
 * Purpose:
 *   各 module page の code block 横に「SJIS でコピー（実験）」ボタンを追加し、
 *   3 つの案 (A/B/C) で clipboard 書込挙動を試す。
 *   どの方式でメモ帳 paste → ANSI 保存 / UTF-8 保存時に CP932 bytes が得られるかを
 *   実機（Windows メモ帳）で検証する PoC。
 *
 * 技術的事実 (起点):
 *   - Chromium の navigator.clipboard.write は ClipboardItem の text/plain Blob を
 *     必ず UTF-8 として decode する (W3C Clipboard API §5)。charset hint は無視。
 *   - Windows clipboard format は CF_UNICODETEXT (UTF-16) で書かれる。
 *   - メモ帳 paste は CF_UNICODETEXT を読む → 保存時のエンコーディング選択で
 *     bytes 形式が決まる。
 *
 *   結論として「SJIS bytes を clipboard に直接書く」のはブラウザ仕様上不可能。
 *   ただし以下 3 案で「メモ帳 UTF-8 保存でも CP932 bytes に落ちる」可能性を検証する。
 */

(function () {
  "use strict";

  // ------------------------------------------------------------
  // Unicode -> CP932 / Shift_JIS 変換用の最小実装
  // ------------------------------------------------------------
  //
  // TextEncoder は UTF-8 専用 (W3C Encoding §10.1)。
  // CP932 への encode table を full に持つと ~100KB 必要。
  // PoC では fetch で同一 module の md (CP932) を取得して
  // code block を bytes 単位で抽出する方式を採用 (encoding.js 不要)。
  //
  // raw URL: https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/docs/<path>
  // GitHub raw は CP932 bytes をそのまま octet-stream で配信する。
  // ------------------------------------------------------------

  var RAW_BASE = "https://raw.githubusercontent.com/ai-crafted-portfolio/knowledgevba/main/docs/";

  function currentMdPath() {
    // window.location.pathname 例: /knowledgevba/modules/admin/thisworkbook.cls/
    // GitHub raw に変換: modules/admin/thisworkbook.cls.md
    var p = window.location.pathname;
    var idx = p.indexOf("/knowledgevba/");
    if (idx < 0) return null;
    var rel = p.slice(idx + "/knowledgevba/".length);
    if (rel.endsWith("/")) rel = rel.slice(0, -1);
    if (!rel) return null;
    return rel + ".md";
  }

  // raw bytes (Uint8Array) から code block (```vb ... ```) を抽出
  function extractCodeBlockBytes(bytes) {
    // backtick = 0x60
    var marker = [0x60, 0x60, 0x60];
    function indexOf3(start) {
      for (var i = start; i <= bytes.length - 3; i++) {
        if (bytes[i] === 0x60 && bytes[i+1] === 0x60 && bytes[i+2] === 0x60) return i;
      }
      return -1;
    }
    var open = indexOf3(0);
    if (open < 0) return null;
    // skip language tag and first \n
    var nl = open + 3;
    while (nl < bytes.length && bytes[nl] !== 0x0a) nl++;
    var bodyStart = nl + 1;
    var close = indexOf3(bodyStart);
    if (close < 0) return null;
    // body end: trim trailing \r\n before ```
    var bodyEnd = close;
    if (bodyEnd > bodyStart && bytes[bodyEnd - 1] === 0x0a) bodyEnd--;
    if (bodyEnd > bodyStart && bytes[bodyEnd - 1] === 0x0d) bodyEnd--;
    return bytes.slice(bodyStart, bodyEnd);
  }

  // --- toast UI ---
  function toast(msg, isError) {
    var el = document.createElement("div");
    el.textContent = msg;
    el.style.cssText = [
      "position:fixed",
      "bottom:20px",
      "right:20px",
      "z-index:9999",
      "max-width:480px",
      "padding:10px 16px",
      "border-radius:6px",
      "font-size:13px",
      "line-height:1.5",
      "box-shadow:0 4px 16px rgba(0,0,0,0.2)",
      "background:" + (isError ? "#c62828" : "#1565c0"),
      "color:#fff",
      "font-family:'Noto Sans JP',sans-serif"
    ].join(";");
    document.body.appendChild(el);
    setTimeout(function () { el.remove(); }, 5000);
  }

  // --- 案 A: charset=shift_jis Blob (charset hint 検証) ---
  async function copyMethodA(sjisBytes) {
    var blob = new Blob([sjisBytes], { type: "text/plain;charset=shift_jis" });
    var item = new ClipboardItem({ "text/plain": blob });
    await navigator.clipboard.write([item]);
  }

  // --- 案 B: latin-1 風 JS string (bytes 0x00-0xFF を U+0000-U+00FF にマップ) ---
  async function copyMethodB(sjisBytes) {
    var s = "";
    for (var i = 0; i < sjisBytes.length; i++) {
      s += String.fromCharCode(sjisBytes[i]);
    }
    await navigator.clipboard.writeText(s);
  }

  // --- 案 C: SJIS -> Unicode で正しく decode した文字列を writeText (基準実装) ---
  async function copyMethodC(sjisBytes) {
    var decoder = new TextDecoder("shift_jis");
    var text = decoder.decode(sjisBytes);
    await navigator.clipboard.writeText(text);
  }

  // 共通: code block bytes 取得 → 各方式実行
  async function runMethod(method, label) {
    var mdPath = currentMdPath();
    if (!mdPath) {
      toast("md path 解決失敗: " + window.location.pathname, true);
      return;
    }
    try {
      var resp = await fetch(RAW_BASE + mdPath, { cache: "no-cache" });
      if (!resp.ok) throw new Error("HTTP " + resp.status);
      var buf = await resp.arrayBuffer();
      var bytes = new Uint8Array(buf);
      var code = extractCodeBlockBytes(bytes);
      if (!code) throw new Error("code block 抽出失敗");
      await method(code);
      toast(
        "[" + label + "] clipboard 書込 OK (" + code.length + " bytes)。" +
        "メモ帳に貼り付け→ANSI または UTF-8 で保存→bytes 確認してください。"
      );
    } catch (e) {
      toast("[" + label + "] 失敗: " + e.message, true);
    }
  }

  // --- ボタン挿入 ---
  function injectButtons() {
    // 対象: module page (URL に /modules/ を含む) の <pre> 内 code block
    if (window.location.pathname.indexOf("/modules/") < 0) return;
    if (currentMdPath() === null) return;

    var pres = document.querySelectorAll("article pre");
    pres.forEach(function (pre) {
      if (pre.dataset.sjisInjected === "1") return;
      pre.dataset.sjisInjected = "1";

      var bar = document.createElement("div");
      bar.style.cssText = [
        "display:flex",
        "gap:8px",
        "flex-wrap:wrap",
        "margin:8px 0",
        "padding:8px",
        "border:1px dashed #888",
        "border-radius:6px",
        "font-size:12px",
        "background:rgba(0,0,0,0.03)"
      ].join(";");

      var label = document.createElement("span");
      label.textContent = "SJIS copy 実験 (PoC):";
      label.style.cssText = "align-self:center;color:#666;";
      bar.appendChild(label);

      [
        ["案A: charset=shift_jis Blob", function (b) { return copyMethodA(b); }],
        ["案B: latin-1 string writeText", function (b) { return copyMethodB(b); }],
        ["案C: SJIS decode → writeText (基準)", function (b) { return copyMethodC(b); }]
      ].forEach(function (pair) {
        var btn = document.createElement("button");
        btn.type = "button";
        btn.textContent = pair[0];
        btn.style.cssText = [
          "padding:4px 10px",
          "border:1px solid #1565c0",
          "background:#fff",
          "color:#1565c0",
          "border-radius:4px",
          "cursor:pointer",
          "font-size:12px"
        ].join(";");
        btn.addEventListener("click", function () { runMethod(pair[1], pair[0]); });
        bar.appendChild(btn);
      });

      pre.parentNode.insertBefore(bar, pre);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", injectButtons);
  } else {
    injectButtons();
  }
  // mkdocs material は instant navigation で history pushState を使う → 再注入
  if (window.document$ && typeof window.document$.subscribe === "function") {
    window.document$.subscribe(injectButtons);
  }
})();
