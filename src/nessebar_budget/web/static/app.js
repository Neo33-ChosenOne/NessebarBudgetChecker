/* Бюджетен монитор Несебър — client-side sort/search/filter for data tables.
 * Vanilla JS, no build step, no dependencies. Progressive enhancement: the
 * table works (and is fully readable/printable) with JS disabled; this only
 * adds sorting/filtering convenience on top of the plain HTML table. */
(function () {
  "use strict";

  function parseSortValue(cell) {
    var raw = cell.getAttribute("data-sort-value");
    if (raw === null) {
      raw = cell.textContent.trim();
    }
    var num = Number(raw);
    if (raw !== "" && !Number.isNaN(num)) {
      return num;
    }
    return raw.toLowerCase();
  }

  function enableSorting(table) {
    var thead = table.tHead;
    if (!thead) return;
    var headerRow = thead.rows[0];
    var tbody = table.tBodies[0];
    if (!tbody) return;

    Array.prototype.forEach.call(headerRow.cells, function (th, index) {
      if (!th.hasAttribute("data-sort")) return;
      th.setAttribute("tabindex", "0");
      th.setAttribute("title", "Подреди по тази колона");
      th.setAttribute("aria-sort", "none");

      function toggleSort() {
        var current = th.getAttribute("data-sort");
        // Number columns start with the largest value: that is what people
        // look for in a table of amounts.
        var first = th.classList.contains("num") ? "desc" : "asc";
        var direction = current === "none" ? first : current === "asc" ? "desc" : "asc";

        Array.prototype.forEach.call(headerRow.cells, function (otherTh) {
          if (otherTh !== th && otherTh.hasAttribute("data-sort")) {
            otherTh.setAttribute("data-sort", "none");
            otherTh.setAttribute("aria-sort", "none");
          }
        });
        th.setAttribute("data-sort", direction);
        th.setAttribute("aria-sort", direction === "asc" ? "ascending" : "descending");

        var rows = Array.prototype.slice.call(tbody.rows);
        rows.sort(function (rowA, rowB) {
          var a = parseSortValue(rowA.cells[index]);
          var b = parseSortValue(rowB.cells[index]);
          if (a < b) return direction === "asc" ? -1 : 1;
          if (a > b) return direction === "asc" ? 1 : -1;
          return 0;
        });
        rows.forEach(function (row) {
          tbody.appendChild(row);
        });
      }

      th.addEventListener("click", toggleSort);
      th.addEventListener("keydown", function (event) {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          toggleSort();
        }
      });
    });
  }

  // "Nothing matches" line, shown instead of an empty table or list.
  function emptyNote(after) {
    var note = document.createElement("p");
    note.className = "empty";
    note.hidden = true;
    note.setAttribute("role", "status");
    note.textContent = "Няма записи, които отговарят на търсенето и филтрите. Изчистете търсенето или изберете друг филтър.";
    after.parentNode.insertBefore(note, after.nextSibling);
    return note;
  }

  function enableFiltering(table) {
    var wrap = table.closest("[data-table-container]");
    if (!wrap) return;
    var searchInput = wrap.querySelector("[data-table-search]");
    var selects = Array.prototype.slice.call(wrap.querySelectorAll("[data-table-filter]"));
    var tbody = table.tBodies[0];
    var countEl = wrap.querySelector("[data-table-count]");
    if (!tbody) return;
    var rows = Array.prototype.slice.call(tbody.rows);
    var note = emptyNote(table.closest(".table-wrap") || table);

    function applyFilters() {
      var query = searchInput ? searchInput.value.trim().toLowerCase() : "";
      var visible = 0;
      rows.forEach(function (row) {
        var matchesSearch = !query || row.textContent.toLowerCase().indexOf(query) !== -1;
        var matchesFilters = selects.every(function (select) {
          if (!select.value) return true;
          return row.getAttribute("data-" + select.getAttribute("data-table-filter")) === select.value;
        });
        var show = matchesSearch && matchesFilters;
        row.hidden = !show;
        // data-no-count: a totals-style row that is not one of the things counted
        if (show && !row.hasAttribute("data-no-count")) visible += 1;
      });
      if (countEl) {
        countEl.textContent = visible.toLocaleString("en-US");
      }
      note.hidden = visible !== 0;
    }

    if (searchInput) {
      searchInput.addEventListener("input", applyFilters);
    }
    selects.forEach(function (select) {
      select.addEventListener("change", applyFilters);
    });
    applyFilters();
  }

  // flags/index.html's flag list: a <details>/<summary> row per flag (so
  // "expand for the full explanation" needs no JS at all) with a search box,
  // a tier segmented control, and severity/rule <select>s filtering which
  // rows are hidden. Structurally similar to enableFiltering() above, but
  // over [data-flag-row] elements instead of <table> rows.
  function enableFlagFilters(container) {
    var rows = Array.prototype.slice.call(container.querySelectorAll("[data-flag-row]"));
    var searchInput = container.querySelector("[data-table-search]");
    var selects = Array.prototype.slice.call(container.querySelectorAll("[data-table-filter]"));
    var tierButtons = Array.prototype.slice.call(container.querySelectorAll("[data-tier-filter]"));
    var countEl = container.querySelector("[data-table-count]");
    var activeTier = "";
    if (!rows.length) return;
    var note = emptyNote(rows[0].parentNode);

    function applyFilters() {
      var query = searchInput ? searchInput.value.trim().toLowerCase() : "";
      var visible = 0;
      rows.forEach(function (row) {
        var matchesSearch = !query || row.textContent.toLowerCase().indexOf(query) !== -1;
        var matchesSelects = selects.every(function (select) {
          if (!select.value) return true;
          return row.getAttribute("data-" + select.getAttribute("data-table-filter")) === select.value;
        });
        var matchesTier = !activeTier || row.getAttribute("data-tier") === activeTier;
        var show = matchesSearch && matchesSelects && matchesTier;
        row.hidden = !show;
        if (show) visible += 1;
      });
      if (countEl) countEl.textContent = visible.toLocaleString("en-US");
      note.hidden = visible !== 0;
    }

    if (searchInput) searchInput.addEventListener("input", applyFilters);
    selects.forEach(function (select) {
      select.addEventListener("change", applyFilters);
    });
    tierButtons.forEach(function (btn) {
      btn.addEventListener("click", function () {
        activeTier = btn.getAttribute("data-tier-filter") || "";
        tierButtons.forEach(function (b) {
          b.setAttribute("aria-pressed", String(b === btn));
        });
        applyFilters();
      });
    });
    applyFilters();
  }

  // Generic "copy this text" button: copies the nearest [data-copy-source]'s
  // text (inside the same <details>) to the clipboard, with a textarea
  // fallback for browsers/contexts without the async Clipboard API.
  function enableCopyButtons() {
    Array.prototype.forEach.call(document.querySelectorAll("[data-copy-button]"), function (btn) {
      btn.addEventListener("click", function () {
        var host = btn.closest("details") || btn.parentElement;
        var source = host ? host.querySelector("[data-copy-source]") : null;
        if (!source) return;
        var text = source.textContent;
        var original = btn.textContent;
        function showCopied() {
          btn.textContent = "Копирано ✓";
          setTimeout(function () {
            btn.textContent = original;
          }, 1800);
        }
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(showCopied, showCopied);
          return;
        }
        var ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        try {
          document.execCommand("copy");
        } catch (e) {
          /* no-op: clipboard just won't be populated */
        }
        document.body.removeChild(ta);
        showCopied();
      });
    });
  }

  // Home page search across the whole site: places, companies, contracts.
  // The index (data/search.json, one small row each) is fetched on first use.
  function enableSiteSearch(form) {
    var input = form.querySelector("input");
    var out = form.querySelector("[data-site-search-results]");
    var root = form.getAttribute("data-root") || "";
    var index = null;
    var loading = false;
    var LIMIT = 12;

    function say(text) {
      var p = document.createElement("p");
      p.className = "sitesearch__count";
      p.textContent = text;
      out.appendChild(p);
    }

    function render() {
      var words = input.value.trim().toLowerCase().split(/\s+/).filter(Boolean);
      out.textContent = "";
      if (!words.length || !index) return;
      var hits = index.filter(function (entry) {
        return words.every(function (word) {
          return entry.h.indexOf(word) !== -1;
        });
      });
      if (!hits.length) {
        say("Няма резултати. Опитайте с част от името или с друга дума.");
        return;
      }
      say(
        hits.length > LIMIT
          ? "Показани са първите " + LIMIT + " от " + hits.length + " резултата. Добавете още една дума, за да ги стесните."
          : hits.length === 1 ? "1 резултат" : hits.length + " резултата"
      );
      var list = document.createElement("ul");
      hits.slice(0, LIMIT).forEach(function (entry) {
        var item = document.createElement("li");
        var link = document.createElement("a");
        link.href = root + entry.u;
        [["sitesearch__kind", entry.k], ["", entry.t], ["sub", entry.x]].forEach(function (part) {
          if (!part[1]) return;
          var span = document.createElement("span");
          if (part[0]) span.className = part[0];
          span.textContent = part[1];
          link.appendChild(span);
        });
        item.appendChild(link);
        list.appendChild(item);
      });
      out.appendChild(list);
    }

    function load() {
      if (index || loading) return;
      loading = true;
      fetch(form.getAttribute("data-site-search"))
        .then(function (response) {
          return response.json();
        })
        .then(function (rows) {
          rows.forEach(function (entry) {
            entry.h = (entry.t + " " + (entry.x || "")).toLowerCase();
          });
          index = rows;
          render();
        })
        .catch(function () {
          loading = false;
          out.textContent = "";
          say("Търсенето не можа да се зареди. Презаредете страницата и опитайте отново.");
        });
    }

    input.addEventListener("focus", load);
    input.addEventListener("input", function () {
      load();
      render();
    });
    form.addEventListener("submit", function (event) {
      event.preventDefault();
    });
  }

  // "/" focuses the page's search field (as in most keyboard-first apps),
  // unless the user is already typing in a form control.
  function enableSearchShortcut() {
    var search = document.querySelector("[data-table-search], [data-site-search] input");
    if (!search) return;
    document.addEventListener("keydown", function (event) {
      if (event.key !== "/" || event.metaKey || event.ctrlKey || event.altKey) return;
      var tag = (document.activeElement && document.activeElement.tagName) || "";
      if (tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA") return;
      event.preventDefault();
      search.focus();
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    Array.prototype.forEach.call(document.querySelectorAll("table[data-sortable]"), enableSorting);
    Array.prototype.forEach.call(document.querySelectorAll("table[data-filterable]"), enableFiltering);
    Array.prototype.forEach.call(document.querySelectorAll("[data-flag-filters]"), enableFlagFilters);
    Array.prototype.forEach.call(document.querySelectorAll("[data-site-search]"), enableSiteSearch);
    enableCopyButtons();
    enableSearchShortcut();
  });
})();
