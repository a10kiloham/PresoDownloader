// Shared status overlay used by both the DocSend and SharePoint download flows.
// Renders a small fixed panel in the bottom-right corner of the page inside a
// shadow root so page CSS cannot restyle it. Safe to inject more than once.
window.DeckOverlay = window.DeckOverlay || (() => {
    const HOST_ID = "deck-downloader-overlay-host";
    const CSS = `
        :host { all: initial; }
        .panel {
            position: fixed; right: 20px; bottom: 20px; z-index: 2147483647;
            display: flex; align-items: center; gap: 12px;
            min-width: 240px; max-width: 380px; padding: 12px 16px;
            background: rgba(24, 24, 27, 0.94); color: #fafafa;
            font: 13px/1.4 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            border-left: 4px solid #5b8be1; border-radius: 10px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
        }
        .panel[hidden] { display: none; }
        .panel.error { border-left-color: #e5484d; }
        .panel.success { border-left-color: #30a46c; }
        .spinner {
            flex: none; width: 18px; height: 18px; border-radius: 50%;
            border: 3px solid rgba(255, 255, 255, 0.25); border-top-color: #5b8be1;
            animation: spin 0.9s linear infinite;
        }
        .spinner[hidden] { display: none; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .body { flex: 1; min-width: 0; }
        .msg { word-wrap: break-word; }
        .bar { margin-top: 8px; height: 5px; border-radius: 3px; background: rgba(255, 255, 255, 0.18); overflow: hidden; }
        .bar[hidden] { display: none; }
        .fill { height: 100%; width: 0; background: #5b8be1; transition: width 0.15s ease-out; }
    `;

    let host = null;
    let els = null;
    let hideTimer = null;

    const ensure = () => {
        if (host && host.isConnected) {
            return;
        }
        host = document.createElement("div");
        host.id = HOST_ID;
        const root = host.attachShadow({ mode: "open" });

        // Prefer constructed stylesheets: they are not subject to the page's CSP.
        try {
            const sheet = new CSSStyleSheet();
            sheet.replaceSync(CSS);
            root.adoptedStyleSheets = [sheet];
        } catch {
            const style = document.createElement("style");
            style.textContent = CSS;
            root.appendChild(style);
        }

        const panel = document.createElement("div");
        panel.className = "panel";
        panel.hidden = true;
        const spinner = document.createElement("div");
        spinner.className = "spinner";
        const body = document.createElement("div");
        body.className = "body";
        const msg = document.createElement("div");
        msg.className = "msg";
        const bar = document.createElement("div");
        bar.className = "bar";
        bar.hidden = true;
        const fill = document.createElement("div");
        fill.className = "fill";

        bar.appendChild(fill);
        body.appendChild(msg);
        body.appendChild(bar);
        panel.appendChild(spinner);
        panel.appendChild(body);
        root.appendChild(panel);
        (document.body || document.documentElement).appendChild(host);

        els = { panel, spinner, msg, bar, fill };
    };

    // show(message, { spinner: true, progress: 0..1, kind: "info" | "error" | "success" })
    const show = (message, opts = {}) => {
        ensure();
        clearTimeout(hideTimer);
        hideTimer = null;
        els.panel.hidden = false;
        els.panel.className = "panel" + (opts.kind ? " " + opts.kind : "");
        els.spinner.hidden = opts.spinner === false;
        els.msg.textContent = message;
        if (typeof opts.progress === "number" && !Number.isNaN(opts.progress)) {
            els.bar.hidden = false;
            els.fill.style.width = Math.round(Math.max(0, Math.min(1, opts.progress)) * 100) + "%";
        } else {
            els.bar.hidden = true;
        }
    };

    const hide = () => {
        clearTimeout(hideTimer);
        hideTimer = null;
        if (els) {
            els.panel.hidden = true;
        }
    };

    // Show a message for `ms` milliseconds, then hide it.
    const flash = (message, ms = 5000, opts = {}) => {
        show(message, { spinner: false, ...opts });
        hideTimer = setTimeout(hide, ms);
    };

    return { show, hide, flash };
})();
