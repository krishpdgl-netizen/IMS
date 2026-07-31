/* app.js
   Shared helpers used by index.html, dashboard.html, assistant.html, and
   reports.html. Mirrors the pattern from the Finance Management System's
   app.js, adapted for inventory (quantities instead of money, no growth
   pills needed for stock counts). */

const API_BASE = window.API_BASE || ""; // same-origin by default; set window.API_BASE before this script loads to point elsewhere

function requireLogin() {
    const name = localStorage.getItem("userName");
    if (!name) {
        window.location.href = "index.html";
        return "";
    }
    return name;
}

function logout() {
    localStorage.removeItem("userName");
    window.location.href = "index.html";
}

async function apiGet(path) {
    const res = await fetch(API_BASE + path);
    if (!res.ok) {
        let detail = res.statusText;
        try { detail = (await res.json()).detail || detail; } catch (e) {}
        throw new Error(detail);
    }
    return res.json();
}

async function apiPost(path, body) {
    const res = await fetch(API_BASE + path, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
    });
    if (!res.ok) {
        let detail = res.statusText;
        try { detail = (await res.json()).detail || detail; } catch (e) {}
        throw new Error(detail);
    }
    return res.json();
}

function fmtQty(v) {
    v = Number(v || 0);
    return v.toLocaleString("en-IN", { maximumFractionDigits: 2 });
}

function agingPill(days, thresholdDays) {
    if (days >= thresholdDays) return `<span class="pill aging">⏳ ${days}d old</span>`;
    return `<span class="pill flat">${days}d old</span>`;
}

let _toastTimer = null;
function showToast(msg) {
    let el = document.getElementById("_toast");
    if (!el) {
        el = document.createElement("div");
        el.id = "_toast";
        el.style.cssText = "position:fixed;bottom:24px;left:50%;transform:translateX(-50%);" +
            "background:#111827;color:#fff;padding:10px 18px;border-radius:8px;font-size:13px;" +
            "z-index:9999;box-shadow:0 4px 16px rgba(0,0,0,.2);max-width:80vw;text-align:center;";
        document.body.appendChild(el);
    }
    el.textContent = msg;
    el.style.display = "block";
    clearTimeout(_toastTimer);
    _toastTimer = setTimeout(() => { el.style.display = "none"; }, 4000);
}
