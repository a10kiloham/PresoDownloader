// SharePoint / OneDrive document downloader.
//
// How SharePoint's built-in PDF viewer gets the file (observed in a HAR capture):
//   1. The page calls RenderListDataAsStream with the EnableMediaTAUrls render
//      option. The response's ListSchema carries a short-lived media token
//      (.driveAccessCode), the regional media host (.mediaBaseUrl), a caller tag
//      (.callerStack) and a URL template (.transformUrl). Each row carries the
//      Graph drive-item URL (.spItemUrl) and a change tag (.ctag).
//   2. It fills the template with method=passthrough and fetches it from the
//      media host with the token in an X-SPOPacToken header. The response body
//      is the original PDF, byte for byte. "Block download" is only a UI policy.
// This script repeats those steps for the file currently open and saves the
// result with a normal browser download.
(() => {
    if (window.__deckDownloaderSharePoint) {
        return;
    }
    window.__deckDownloaderSharePoint = true;

    const O = window.DeckOverlay;
    const FLASH_MS = 5000;
    const UNREACHABLE = "File Unreachable";
    // Office formats the media service can convert to PDF via transform/pdf.
    const CONVERTIBLE = new Set(["doc", "docx", "dot", "dotx", "ppt", "pptx", "pps", "ppsx", "pot", "potx", "xls", "xlsx", "odt", "odp", "ods", "rtf"]);
    // Same RenderOptions the viewer uses: ContextInfo | ListData | ListSchema | EnableMediaTAUrls | 0x100000
    const RENDER_OPTIONS = 1052679;

    let running = false;

    class Unreachable extends Error {}

    const log = (...args) => console.log("[DeckDownloader]", ...args);

    const fmtBytes = (n) => (n >= 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB");

    const xmlEscape = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");

    // OData string literal: single quotes doubled, then percent-encoded.
    const odataStr = (s) => encodeURIComponent(s.replace(/'/g, "''"));

    const extOf = (name) => {
        const m = /\.([a-z0-9]+)$/i.exec(name || "");
        return m ? m[1].toLowerCase() : "";
    };

    // ---- Locate the document the user is looking at ---------------------------

    const getTargetPath = () => {
        const url = new URL(location.href);
        let id = url.searchParams.get("id") || url.searchParams.get("file");
        if (!id) {
            // Direct file URLs like /sites/x/Shared Documents/deck.pdf?web=1
            const path = decodeURIComponent(url.pathname);
            if (extOf(path) && !/\.aspx$/i.test(path)) {
                id = path;
            }
        }
        if (!id) {
            return null;
        }
        id = id.replace(/\/+$/, "");
        return id.startsWith("/") ? id : "/" + id;
    };

    // ---- SharePoint REST helpers ----------------------------------------------

    const spGet = async (url) => {
        const resp = await fetch(url, { headers: { Accept: "application/json;odata=nometadata" }, credentials: "include" });
        if (!resp.ok) {
            throw new Unreachable(`GET ${url} -> HTTP ${resp.status}`);
        }
        return resp.json();
    };

    const getDigest = async (webUrl) => {
        const resp = await fetch(`${webUrl}/_api/contextinfo`, {
            method: "POST",
            headers: { Accept: "application/json;odata=nometadata" },
            credentials: "include"
        });
        if (!resp.ok) {
            throw new Unreachable(`contextinfo -> HTTP ${resp.status}`);
        }
        const json = await resp.json();
        return json.FormDigestValue || json?.d?.GetContextWebInformation?.FormDigestValue;
    };

    // Web URL: from _spPageContextInfo when the service worker could read it,
    // otherwise probe upwards from the file path until /_api/web answers.
    const resolveWebUrl = async (ctx, filePath) => {
        if (ctx?.webAbsoluteUrl) {
            return ctx.webAbsoluteUrl.replace(/\/+$/, "");
        }
        const segs = filePath.split("/").filter(Boolean);
        for (let n = segs.length - 1; n >= 0; n--) {
            const candidate = location.origin + (n ? "/" + segs.slice(0, n).map(encodeURIComponent).join("/") : "");
            try {
                const json = await spGet(`${candidate}/_api/web?$select=ServerRelativeUrl`);
                if (json && "ServerRelativeUrl" in json) {
                    return candidate;
                }
            } catch {
                // not a web, keep walking up
            }
        }
        throw new Unreachable("Could not resolve the SharePoint web for " + filePath);
    };

    const fileApi = (webUrl, filePath) => `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl=@f)`;
    const fileArg = (filePath) => `@f='${odataStr(filePath)}'`;

    // Returns { id, name } for the list item behind the file. 404 means unreachable.
    const getListItem = async (webUrl, filePath) => {
        const json = await spGet(`${fileApi(webUrl, filePath)}/ListItemAllFields?${fileArg(filePath)}&$select=Id,FileLeafRef`);
        if (!json || json.Id === undefined) {
            throw new Unreachable("No list item for " + filePath);
        }
        return { id: json.Id, name: json.FileLeafRef };
    };

    const resolveListUrl = async (ctx, webUrl, filePath) => {
        if (ctx?.listUrl && filePath.toLowerCase().startsWith(ctx.listUrl.toLowerCase().replace(/\/+$/, "") + "/")) {
            return ctx.listUrl.replace(/\/+$/, "");
        }
        const json = await spGet(`${fileApi(webUrl, filePath)}/ListItemAllFields/ParentList/RootFolder?${fileArg(filePath)}&$select=ServerRelativeUrl`);
        if (!json?.ServerRelativeUrl) {
            throw new Unreachable("Could not resolve the document library for " + filePath);
        }
        return json.ServerRelativeUrl;
    };

    // Same call the viewer makes: returns list schema (media host + token) and the item row.
    const renderListData = async (webUrl, listUrl, itemId, digest) => {
        const url = `${webUrl}/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1='${odataStr(listUrl)}'`;
        const viewXml =
            `<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="ID"/><Value Type="Number">${itemId}</Value></Eq></Where></Query>` +
            `<ViewFields><FieldRef Name="File_x0020_Type"/><FieldRef Name="FileRef"/><FieldRef Name="FileLeafRef"/></ViewFields>` +
            `<RowLimit Paged="TRUE">1</RowLimit></View>`;
        const body = {
            parameters: {
                __metadata: { type: "SP.RenderListDataParameters" },
                RenderOptions: RENDER_OPTIONS,
                ViewXml: viewXml,
                AddRequiredFields: true
            }
        };
        const resp = await fetch(url, {
            method: "POST",
            credentials: "include",
            headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest
            },
            body: JSON.stringify(body)
        });
        if (!resp.ok) {
            throw new Unreachable(`RenderListDataAsStream -> HTTP ${resp.status}`);
        }
        const json = await resp.json();
        const schema = json.ListSchema || {};
        const row = json.ListData?.Row?.[0];
        if (!row || !row[".spItemUrl"]) {
            throw new Unreachable("RenderListDataAsStream returned no row for item " + itemId);
        }
        if (!schema[".driveAccessCode"] || !schema[".mediaBaseUrl"]) {
            throw new Unreachable("RenderListDataAsStream response lacks media token or media host");
        }
        return { schema, row };
    };

    // Fill the viewer's URL template. method is "passthrough" (raw bytes) or "pdf" (convert to PDF).
    const buildMediaUrl = (schema, row, method, fileType) => {
        let template = schema[".transformUrl"];
        if (!template && schema[".pdfConversionUrl"]) {
            template = schema[".pdfConversionUrl"].replace("/transform/pdf?", "/transform/{.method}?");
        }
        if (!template) {
            throw new Unreachable("No media URL template in list schema");
        }
        return template
            .replace("{.mediaBaseUrl}", schema[".mediaBaseUrl"])
            .replace("{.method}", method)
            .replace("{.fileType}", fileType)
            .replace("{.callerStack}", schema[".callerStack"] || "")
            .replace("{.spItemUrl}", encodeURIComponent(row[".spItemUrl"]))
            .replace("{.driveAccessToken}", "cTag=" + encodeURIComponent(row[".ctag"] || ""));
    };

    // Download with progress reporting. Falls back to the access_token query
    // parameter if the header form is rejected.
    const downloadBytes = async (url, schema, onProgress) => {
        const headers = { "X-SPOPacToken": schema[".driveAccessCode"] };
        let resp = await fetch(url, { headers });
        if ((resp.status === 401 || resp.status === 403) && schema[".driveAccessToken"]) {
            log("Header auth rejected with", resp.status, "- retrying with access_token query parameter");
            resp = await fetch(url + "&" + schema[".driveAccessToken"], { headers });
        }
        if (!resp.ok) {
            throw new Unreachable(`media service -> HTTP ${resp.status}`);
        }
        const total = Number(resp.headers.get("content-length")) || 0;
        const reader = resp.body.getReader();
        const chunks = [];
        let received = 0;
        for (;;) {
            const { done, value } = await reader.read();
            if (done) {
                break;
            }
            chunks.push(value);
            received += value.byteLength;
            onProgress(received, total);
        }
        const out = new Uint8Array(received);
        let offset = 0;
        for (const chunk of chunks) {
            out.set(chunk, offset);
            offset += chunk.byteLength;
        }
        return out;
    };

    const looksLikePdf = (bytes) => {
        const head = new TextDecoder("latin1").decode(bytes.subarray(0, 1024));
        return head.includes("%PDF-");
    };

    const saveBlob = (bytes, filename) => {
        const blob = new Blob([bytes], { type: "application/pdf" });
        const blobUrl = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = blobUrl;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
    };

    // ---- Main job -------------------------------------------------------------

    const run = async (ctx) => {
        const filePath = getTargetPath();
        if (!filePath) {
            O.flash("Open a document first, then click the extension.", FLASH_MS, { kind: "error" });
            return;
        }
        log("Target file:", filePath);
        O.show("Locating file…");

        const webUrl = await resolveWebUrl(ctx, filePath);
        const item = await getListItem(webUrl, filePath);
        const listUrl = await resolveListUrl(ctx, webUrl, filePath);
        log("web:", webUrl, "list:", listUrl, "item:", item);

        O.show("Requesting access token…");
        const digest = await getDigest(webUrl);
        const { schema, row } = await renderListData(webUrl, listUrl, item.id, digest);

        const name = row.FileLeafRef || item.name || filePath.split("/").pop();
        const fileType = (row[".fileType"] || row.File_x0020_Type || extOf(name)).toLowerCase();
        let method;
        let outName;
        if (fileType === "pdf") {
            method = "passthrough";
            outName = name;
        } else if (CONVERTIBLE.has(fileType)) {
            method = "pdf";
            outName = name.replace(/\.[^.]+$/, "") + ".pdf";
        } else {
            O.flash(`Unsupported file type: .${fileType || "?"}`, FLASH_MS, { kind: "error" });
            return;
        }

        const mediaUrl = buildMediaUrl(schema, row, method, fileType);
        log("Media URL:", mediaUrl);
        O.show(method === "pdf" ? "Converting to PDF…" : "Downloading…", { progress: 0 });
        const started = Date.now();
        const bytes = await downloadBytes(mediaUrl, schema, (received, total) => {
            if (total) {
                O.show(`Downloading… ${fmtBytes(received)} / ${fmtBytes(total)}`, { progress: received / total });
            } else {
                O.show(`Downloading… ${fmtBytes(received)}`);
            }
        });
        if (!bytes.byteLength || !looksLikePdf(bytes)) {
            throw new Unreachable("Media service response is not a PDF");
        }

        saveBlob(bytes, outName);
        O.flash(`Saved ${outName} (${fmtBytes(bytes.byteLength)} in ${((Date.now() - started) / 1000).toFixed(1)} s)`, FLASH_MS, { kind: "success" });
    };

    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        if (message?.requestType !== "GENERATE_PDF") {
            return;
        }
        if (running) {
            O.show("Please wait, download still in progress…");
            sendResponse({ status: "busy" });
            return;
        }
        running = true;
        sendResponse({ status: "started" });
        run(message.pageContext)
            .catch((err) => {
                console.error("[DeckDownloader]", err);
                O.flash(UNREACHABLE, FLASH_MS, { kind: "error" });
            })
            .finally(() => {
                running = false;
            });
    });
})();
