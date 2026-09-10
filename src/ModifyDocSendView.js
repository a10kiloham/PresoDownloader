// DocSend page helpers. Status messages go through the shared DeckOverlay so the
// UI no longer depends on DocSend's own alert markup.
var showCustomAlert = (message, progress) => DeckOverlay.show(message, { progress });
var showDefaultAlert = (message, kind) => DeckOverlay.flash(message, 5000, { kind: kind || "info" });
var hideCustomAlert = () => DeckOverlay.hide();
var hideDefaultAlert = () => DeckOverlay.hide();

// Strip filesystem-unsafe characters and surrounding whitespace/quotes.
var sanitizeFilename = (name) =>
    name.replace(/^["'\s]+|["'\s]+$/g, "").replace(/[\\/:*?"<>|]/g, "").trim() || "slidedeck";

// Deck name: contact-card description, then sender email domain, then the tab title.
var getDeckName = () => {
    const description = document.getElementsByClassName("contact-card_description")[0];
    if (description && description.innerText.trim()) {
        return sanitizeFilename(description.innerText) + " Deck";
    }
    const email = document.getElementsByClassName("contact-card_email")[0];
    if (email && email.href && email.href.includes("@")) {
        const domain = email.href.split("@")[1].split(".")[0];
        if (domain) {
            return sanitizeFilename(domain) + " Deck";
        }
    }
    const title = document.title.replace(/\s*[-|]\s*DocSend.*$/i, "").trim();
    if (title) {
        return sanitizeFilename(title);
    }
    return "slidedeck";
};

var initiateDownload = (url, filename) => {
    const downloadLink = document.createElement("a");
    downloadLink.href = url;
    downloadLink.setAttribute("download", `${filename}.pdf`);
    downloadLink.click();
};
