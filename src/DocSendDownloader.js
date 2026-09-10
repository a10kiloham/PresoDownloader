// DocSend slide deck downloader: reads each slide's image URL from DocSend's
// page_data endpoint, then assembles the images into a PDF.
(() => {
    if (window.__deckDownloaderDocSend) {
        return;
    }
    window.__deckDownloaderDocSend = true;

    const FLASH_MS = 5000;
    let running = false;

    const userIsAuthenticated = () => document.getElementById("prompt") == null;

    const getNumSlides = () => {
        const label = document.getElementsByClassName("page-label")[0];
        const n = label ? parseInt(label.innerHTML.split(" ")[0], 10) : NaN;
        return Number.isFinite(n) && n > 0 ? n : null;
    };

    // Returns one loader per slide. Each loader asks DocSend for that slide's
    // signed image URL only when the slide is about to be rendered, because the
    // URLs expire and long decks fail with AccessDenied if fetched up front.
    const getSlideLoaders = (numSlides) => {
        const baseUrl = window.location.href.split("?")[0].replace(/\/+$/, "");
        return Array.from({ length: numSlides }, (_, i) => async () => {
            const response = await fetchWithRetry(`${baseUrl}/page_data/${i + 1}`);
            const data = await response.json();
            if (!data.imageUrl) {
                throw new Error(`page_data/${i + 1} has no imageUrl`);
            }
            return data.imageUrl;
        });
    };

    const run = async () => {
        if (!userIsAuthenticated()) {
            showDefaultAlert("You must be signed in to download this slide deck as a PDF.", "error");
            return;
        }
        const numSlides = getNumSlides();
        if (!numSlides) {
            throw new Error("Could not read the slide count from the page");
        }
        const deckName = getDeckName();
        showCustomAlert(`Generating slide deck as PDF: 0/${numSlides} slides complete…`, 0);
        const startTime = Date.now();
        const blobUrl = await buildPdf(getSlideLoaders(numSlides), deckName, (done, total) => {
            showCustomAlert(`Generating slide deck as PDF: ${done}/${total} slides complete…`, done / total);
        });
        initiateDownload(blobUrl, deckName);
        hideCustomAlert();
        showDefaultAlert(`Done! Slide deck PDF generated in ${((Date.now() - startTime) / 1000).toFixed(1)} s.`, "success");
    };

    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
        if (message?.requestType !== "GENERATE_PDF") {
            return;
        }
        if (running) {
            showCustomAlert("Please wait. Still generating slide deck as PDF…");
            sendResponse({ status: "busy" });
            return;
        }
        running = true;
        sendResponse({ status: "started" });
        run()
            .catch((err) => {
                console.error("[DeckDownloader]", err);
                hideCustomAlert();
                showDefaultAlert("File Unreachable", "error");
            })
            .finally(() => {
                running = false;
            });
    });
})();
