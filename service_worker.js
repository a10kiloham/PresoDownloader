// Routes the toolbar click to the right downloader for the current tab.
const DOCSEND_SCRIPTS = [
    "./modules/pdfkit.js",
    "./modules/blob-stream.js",
    "./src/Overlay.js",
    "./src/ModifyDocSendView.js",
    "./src/GeneratePDF.js",
    "./src/DocSendDownloader.js"
];
const SHAREPOINT_SCRIPTS = ["./src/Overlay.js", "./src/SharePointDownloader.js"];

const detectSite = (url) => {
    let host;
    try {
        host = new URL(url).hostname.toLowerCase();
    } catch {
        return null;
    }
    if (host === "docsend.com" || host.endsWith(".docsend.com") || host === "docsend.dropbox.com" || host.endsWith(".docsend.dropbox.com")) {
        return "docsend";
    }
    if (host.endsWith(".sharepoint.com")) {
        return "sharepoint";
    }
    return null;
};

// Read _spPageContextInfo from the page's own JS world (content scripts cannot see it).
const readSharePointContext = async (tabId) => {
    try {
        const [res] = await chrome.scripting.executeScript({
            target: { tabId },
            world: "MAIN",
            func: () => {
                const c = window._spPageContextInfo;
                if (!c) {
                    return null;
                }
                return {
                    webAbsoluteUrl: c.webAbsoluteUrl,
                    webServerRelativeUrl: c.webServerRelativeUrl,
                    listUrl: c.listUrl,
                    listId: c.listId
                };
            }
        });
        return res?.result || null;
    } catch (err) {
        console.warn("Could not read _spPageContextInfo:", err);
        return null;
    }
};

const sendJob = (tabId, message) => chrome.tabs.sendMessage(tabId, message);

const flashBadge = async (tabId, title) => {
    try {
        await chrome.action.setBadgeBackgroundColor({ tabId, color: "#e5484d" });
        await chrome.action.setBadgeText({ tabId, text: "!" });
        await chrome.action.setTitle({ tabId, title });
        setTimeout(() => {
            chrome.action.setBadgeText({ tabId, text: "" }).catch(() => {});
            chrome.action.setTitle({ tabId, title: "Download this document as a PDF" }).catch(() => {});
        }, 5000);
    } catch {
        // tab may be gone
    }
};

chrome.action.onClicked.addListener(async (tab) => {
    const site = detectSite(tab.url || "");
    if (!site) {
        await flashBadge(tab.id, "Open a DocSend deck or a SharePoint document first");
        return;
    }
    try {
        const message = { requestType: "GENERATE_PDF" };
        if (site === "sharepoint") {
            message.pageContext = await readSharePointContext(tab.id);
        }
        try {
            await sendJob(tab.id, message);
        } catch {
            // No content script in this tab yet: inject, then retry.
            await chrome.scripting.executeScript({
                target: { tabId: tab.id },
                files: site === "docsend" ? DOCSEND_SCRIPTS : SHAREPOINT_SCRIPTS
            });
            await sendJob(tab.id, message);
        }
    } catch (err) {
        console.error("Deck downloader failed to start:", err);
        await flashBadge(tab.id, "Could not start the download on this page");
    }
});
