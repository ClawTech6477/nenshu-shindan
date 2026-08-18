/* 記事内アフィリンク (2026-08-18新設)
 *
 * なぜ作ったか: 実測で page_view 271 → diagnosis_start 11 (4.1%) しかなく、
 * 記事を読んだ人の95.9%が診断に触れずに帰っていた。記事は職種別なので、
 * その職種のサービスは既に確定している。診断を挟まず直接送れるようにする。
 *
 * データは blog/affiliates.json（scripts/affiliate/blog_affiliates_build.py が
 * 診断ページの承認済みリンクから毎日生成）。提携が承認された翌日に自動で増える。
 * 未承認の職種は0件になり、その場合はセクションごと出さない（空枠を出さない）。
 */
/* 2026-08-19追加: 本文中のテキストリンク(.affi-inline)
 *
 * なぜ: A8実測で、記事本文の文脈内テキストリンクは imp5/click4(CTR80%)、
 * 一方カード型の一覧枠は imp290/click3(CTR1.0%)だった。同じ表示回数あたり
 * 数十倍違う。記事末尾のカード枠だけでは、本文を読んで離脱する読者を取り逃す。
 *
 * 使い方(記事側HTML):
 *   <span class="affi-inline" data-job="constr" data-rank="1"
 *         data-fallback="建設業界に特化した転職エージェント"></span>
 * 承認済みリンクがあればサービス名のaタグに、無ければ data-fallback の
 * 素のテキストに置き換わる。未承認でも文章が壊れないようにするための設計。
 */
function renderInline(data) {
  var rendered = 0;
  document.querySelectorAll(".affi-inline").forEach(function (el) {
    var job = el.getAttribute("data-job");
    var rank = Number(el.getAttribute("data-rank") || 1);
    var fb = el.getAttribute("data-fallback") || "";
    var list = (data && data.jobs && data.jobs[job]) || [];
    // data-rank は「できればこの順位を」という希望。その順位が無い職種
    // (承認1社だけ等)では先頭に落として、リンクを1本も失わないようにする。
    var s = list[rank - 1] || list[0];
    if (!s) { el.textContent = fb; return; }   // 未承認: 素のテキストで文章を保つ
    var a = document.createElement("a");
    a.href = s.url;
    a.target = "_blank";
    a.rel = "nofollow sponsored noopener";
    a.className = "affi-inline-a";
    a.textContent = s.name;
    a.addEventListener("click", function () {
      if (typeof gtag === "function") {
        gtag("event", "article_affiliate_click", {
          site_name: s.name, rank: rank, job: job,
          placement: "inline", article: location.pathname
        });
      }
    });
    el.textContent = "";
    el.appendChild(a);
    rendered++;
  });
  // 広告表記は「実際にリンクが出た時だけ」出す。未承認で素のテキストに
  // なった記事に広告表記だけ残ると、事実と違う表示になるため。
  if (rendered > 0) {
    var art = document.querySelector("article");
    if (art && !document.querySelector(".affi-note")) {
      var note = document.createElement("p");
      note.className = "affi-note";
      note.textContent = "本文中の一部リンクは広告（アフィリエイト）です。"
        + "リンク先の各社は求職者側の費用が無料で、掲載順は当サイトの計測データに基づきます。";
      var faq = document.getElementById("faq");
      if (faq) { art.insertBefore(note, faq); } else { art.appendChild(note); }
    }
  }
}

(function () {
  // 本文中リンクは末尾カード枠が無い記事にも入れられるよう、独立して取得・描画する
  if (document.querySelector(".affi-inline")) {
    fetch("affiliates.json", { cache: "no-cache" })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(renderInline)
      .catch(function () { renderInline(null); });
  }

  var box = document.querySelector(".job-affiliates");
  if (!box) return;
  var job = box.getAttribute("data-job");
  if (!job) return;

  fetch("affiliates.json", { cache: "no-cache" })
    .then(function (r) { return r.ok ? r.json() : null; })
    .then(function (data) {
      if (!data || !data.jobs) return;
      var list = data.jobs[job];
      if (!list || !list.length) return;   // 承認済みが無い職種は何も出さない

      var h = '<h2>この職種で使える転職サービス</h2>'
            + '<p class="ja-lead">下の各社は求職者側の費用が無料で、料金は採用が決まった時に'
            + '採用側が負担する仕組みです。2〜3社を併用して比較するのが基本形になります。</p>'
            + '<ul class="ja-list">';
      list.forEach(function (s, i) {
        h += '<li><a href="' + s.url + '" target="_blank" rel="nofollow sponsored noopener"'
           + ' data-name="' + s.name + '" data-rank="' + (i + 1) + '">'
           + '<span class="ja-name">' + s.name + '</span>'
           + '<span class="ja-desc">' + s.desc + '</span>'
           + '<span class="ja-go">無料で相談する →</span></a></li>';
      });
      h += '</ul><p class="ja-note">本セクションは広告です。掲載順は当サイトの計測データに'
         + '基づくもので、報酬額の多寡で順位を操作することはありません。</p>';
      box.innerHTML = h;
      box.classList.add("is-ready");

      box.querySelectorAll("a").forEach(function (a) {
        a.addEventListener("click", function () {
          if (typeof gtag === "function") {
            gtag("event", "article_affiliate_click", {
              site_name: a.getAttribute("data-name"),
              rank: Number(a.getAttribute("data-rank")),
              job: job,
              article: location.pathname
            });
          }
        });
      });
    })
    .catch(function () { /* 取得失敗時は何も出さない */ });
})();
