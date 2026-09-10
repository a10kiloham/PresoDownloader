// Builds a PDF from DocSend slides with pdfkit + blob-stream.
// Each slide's signed image URL is requested right before the image is fetched
// (rather than all up front) because the links expire on long decks.
var fetchWithRetry = async (url, attempts = 3) => {
    let lastError;
    for (let i = 0; i < attempts; i++) {
        try {
            const response = await fetch(url);
            if (!response.ok) {
                throw new Error(`HTTP ${response.status} for ${url}`);
            }
            return response;
        } catch (e) {
            lastError = e;
            if (i < attempts - 1) {
                await new Promise((r) => setTimeout(r, 500 * (i + 1)));
            }
        }
    }
    throw lastError;
};

var getImageAsDataUrl = async (url) => {
    const blob = await (await fetchWithRetry(url)).blob();
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
};

// slides: array of async functions, each resolving to the slide's image URL.
var addSlidesToPDF = async (doc, slides, onSlide) => {
    for (let i = 0; i < slides.length; i++) {
        const imageUrl = await slides[i]();
        const data = await getImageAsDataUrl(imageUrl);
        const img = doc.openImage(data);
        doc.addPage({ size: [img.width, img.height] });
        doc.image(img, 0, 0);
        onSlide(i + 1, slides.length);
    }
};

// Resolves with a blob: URL of the finished PDF. A fresh document is created per
// call so the deck can be downloaded again without reloading the page.
var buildPdf = (slides, deckName, onSlide) =>
    new Promise((resolve, reject) => {
        const doc = new PDFDocument({ layout: "landscape", margin: 0, autoFirstPage: false });
        doc.info.Title = deckName;
        doc.info.Creator = "DocSend & SharePoint Deck Downloader";
        doc.info.Producer = "DocSend & SharePoint Deck Downloader";
        const stream = doc.pipe(blobStream());
        stream.on("finish", () => resolve(stream.toBlobURL("application/pdf")));
        stream.on("error", reject);
        addSlidesToPDF(doc, slides, onSlide)
            .then(() => doc.end())
            .catch(reject);
    });
