/* global Office, Word */

let selectedStyle = "ieee";

const STYLES = ["ieee", "apa", "mla", "harvard", "chicago", "vancouver"];

Office.onReady((info) => {
    const formatBtn = document.getElementById("formatBtn");
    const styleSelector = document.getElementById("styleSelector");

    styleSelector.addEventListener("click", (event) => {
        const btn = event.target.closest(".style-btn");

        if (!btn) return;

        styleSelector
            .querySelectorAll(".style-btn")
            .forEach((b) => b.classList.remove("is-active"));

        btn.classList.add("is-active");

        selectedStyle = btn.dataset.style;
    });

    if (info.host === Office.HostType.Word) {
        formatBtn.disabled = false;

        setStatus(
            "info",
            "Ready. Select references in Word to begin."
        );

        formatBtn.onclick = formatSelection;
    } else {
        setStatus(
            "error",
            "This add-in only works inside Microsoft Word."
        );
    }
});


function setStatus(type, message) {
    const statusDiv = document.getElementById("status");

    const icons = {
        info: "ℹ️",
        success: "✅",
        error: "⚠️"
    };

    statusDiv.className = `status status--${type}`;

    statusDiv.innerHTML =
        `<span class="status__icon">${icons[type]}</span>` +
        `<span class="status__text"></span>`;

    statusDiv.querySelector(".status__text").innerText = message;
}


function setLoading(isLoading) {
    const formatBtn = document.getElementById("formatBtn");
    const spinner = document.getElementById("btnSpinner");
    const label = document.getElementById("btnLabel");

    formatBtn.disabled = isLoading;

    spinner.hidden = !isLoading;

    label.innerText = isLoading
        ? "Formatting…"
        : "Format Selected References";
}


async function formatSelection() {
    setLoading(true);

    try {
        await Word.run(async (context) => {

            const selection = context.document.getSelection();

            selection.load("text");

            await context.sync();


            const rawLines = selection.text
                .split(/[\r\v\n]+/)
                .map((line) => line.trim())
                .filter((line) => line.length > 0);


            if (rawLines.length === 0) {
                setStatus(
                    "error",
                    "Please highlight one or more references in your Word document first."
                );

                return;
            }


            const references = rawLines.map((originalText, i) => {

                const index = i + 1;

                const parsedData =
                    parseCitationText(originalText);


                return {
                    index,
                    originalText,
                    parsedData,

                    formattedText:
                        buildStyleString(
                            parsedData,
                            selectedStyle,
                            index
                        )
                };
            });


            /*
             * Single atomic replace for the whole reference list.
             * Doing every line in ONE insertText call avoids Word's
             * document positions shifting while references are updated.
             */

            const joinedText =
                references
                    .map((r) => r.formattedText)
                    .join("\r");


            selection.insertText(
                joinedText,
                Word.InsertLocation.replace
            );


            await context.sync();


            setStatus(
                "info",
                "Reference list updated. Syncing in-text citations…"
            );


            let inlineUpdates = 0;


            for (const ref of references) {

                inlineUpdates +=
                    await syncInlineCitations(
                        context,
                        ref
                    );

            }


            reportSuccess(
                references.length,
                inlineUpdates
            );

        });

    } catch (error) {

        setStatus(
            "error",
            error.message
        );

    } finally {

        setLoading(false);

    }
}


/**
 * Best-effort in-text citation sync.
 *
 * There is no persistent link between a reference-list entry and
 * its in-text marker elsewhere in the document.
 *
 * Therefore this function searches for common citation patterns
 * and replaces them with the marker for the selected style.
 */

async function syncInlineCitations(context, ref) {

    const {
        index,
        parsedData
    } = ref;


    const lastName =
        parsedData.author
            .split(",")[0]
            .trim();


    const newMarker =
        buildInlineMarker(
            parsedData,
            selectedStyle,
            index
        );


    const candidates = new Set([

        `[${index}]`,

        `(${index})`,

        `(${lastName}, ${parsedData.year})`,

        `(${lastName} ${parsedData.year})`,

        `(${lastName})`

    ]);


    candidates.delete(newMarker);


    let totalReplacements = 0;


    for (const candidate of candidates) {

        const searchResults =
            context.document.body.search(
                candidate,
                {
                    matchCase: false
                }
            );


        searchResults.load("items");

        await context.sync();


        if (searchResults.items.length === 0) {
            continue;
        }


        for (const item of searchResults.items) {

            item.insertText(
                newMarker,
                Word.InsertLocation.replace
            );


            await context.sync();


            totalReplacements++;

        }

    }


    return totalReplacements;
}


function buildInlineMarker(
    parsedData,
    style,
    index
) {

    const lastName =
        parsedData.author
            .split(",")[0]
            .trim();


    switch (style) {

        case "ieee":

            return `[${index}]`;


        case "vancouver":

            return `(${index})`;


        case "chicago":

            return `(${lastName} ${parsedData.year})`;


        case "mla":

            return `(${lastName})`;


        case "apa":

        case "harvard":

        default:

            return `(${lastName}, ${parsedData.year})`;
    }
}


function reportSuccess(
    refCount,
    inlineCount
) {

    const refPlural =
        refCount === 1
            ? "reference"
            : "references";


    const inlinePart =
        inlineCount > 0

            ? ` Updated ${inlineCount} in-text citation${inlineCount === 1 ? "" : "s"}.`

            : " No matching in-text citations found to update.";


    setStatus(
        "success",
        `Formatted ${refCount} ${refPlural} to ${selectedStyle.toUpperCase()}.${inlinePart}`
    );
}


/**
 * Patterns for recognizing references already formatted
 * in the six supported styles.
 */

const KNOWN_STYLE_PATTERNS = [

    {
        style: "ieee",

        regex:
            /^\[\d+\]\s*(.+?),\s*"(.+?),"\s*(.+?),\s*(n\.d\.|\d{4})\.$/,

        map: (m) => ({
            author: m[1],
            title: m[2],
            source: m[3],
            year: m[4]
        })
    },


    {
        style: "vancouver",

        regex:
            /^\d+\.\s*(.+?)\.\s*(.+?)\.\s*(.+?)\.\s*(n\.d\.|\d{4})\.$/,

        map: (m) => ({
            author: m[1],
            title: m[2],
            source: m[3],
            year: m[4]
        })
    },


    {
        style: "harvard",

        regex:
            /^(.+?)\s*\((n\.d\.|\d{4})\)\s*'(.+?)',\s*(.+?)\.$/,

        map: (m) => ({
            author: m[1],
            year: m[2],
            title: m[3],
            source: m[4]
        })
    },


    {
        style: "chicago",

        regex:
            /^(.+?)\.\s*(n\.d\.|\d{4})\.\s*"(.+?)"\.\s*(.+?)\.$/,

        map: (m) => ({
            author: m[1],
            year: m[2],
            title: m[3],
            source: m[4]
        })
    },


    {
        style: "mla",

        regex:
            /^(.+?)\.\s*"(.+?)"\.\s*(.+?),\s*(n\.d\.|\d{4})\.$/,

        map: (m) => ({
            author: m[1],
            title: m[2],
            source: m[3],
            year: m[4]
        })
    },


    {
        style: "apa",

        regex:
            /^(.+?)\s*\((n\.d\.|\d{4})\)\.\s*(.+?)\.\s*(.+?)\.$/,

        map: (m) => ({
            author: m[1],
            year: m[2],
            title: m[3],
            source: m[4]
        })
    }

];


function parseCitationText(text) {

    for (
        const { regex, map }
        of KNOWN_STYLE_PATTERNS
    ) {

        const match =
            text.match(regex);


        if (match) {

            const fields =
                map(match);


            return {

                author:
                    fields.author.trim(),

                year:
                    fields.year.trim(),

                title:
                    fields.title.trim(),

                source:
                    fields.source.trim()

            };
        }
    }


    return parseRawCitationText(text);
}


function parseRawCitationText(text) {

    const yearMatch =
        text.match(
            /\b(19|20)\d{2}\b/
        );


    const year =
        yearMatch
            ? yearMatch[0]
            : "n.d.";


    let cleanText =
        text

            .replace(
                /\b(19|20)\d{2}\b/,
                ""
            )

            .replace(
                /\(\s*\)/g,
                ""
            )

            .replace(
                /\[\s*\]/g,
                ""
            )

            .trim();


    let author =
        "Author, A.";


    const authorMatch =
        cleanText.match(
            /^([A-Z][a-z]+(?:,\s*[A-Z]\.|\s+[A-Z]\.)*(?:\s*,?\s*and\s+|\s*,?\s*&\s*)?)+/
        );


    if (
        authorMatch &&
        authorMatch[0].length > 2
    ) {

        author =
            authorMatch[0]
                .trim()
                .replace(
                    /,\s*$/,
                    ""
                );
    }


    let titleAndSource =
        authorMatch

            ? cleanText
                .slice(
                    authorMatch[0].length
                )
                .trim()

            : cleanText;


    titleAndSource =
        titleAndSource

            .replace(
                /^[.,\s]+/,
                ""
            )

            .replace(
                /[.,\s]+$/,
                ""
            );


    let title =
        titleAndSource;


    let source =
        "Publication Source";


    const quotesMatch =
        titleAndSource.match(
            /["'\u201c](.*?)["'\u201d]/
        );


    if (quotesMatch) {

        title =
            quotesMatch[1];


        source =
            titleAndSource

                .replace(
                    quotesMatch[0],
                    ""
                )

                .replace(
                    /^[.,\s]+/,
                    ""
                );

    } else if (
        titleAndSource.includes(".")
    ) {

        const parts =
            titleAndSource.split(".");


        title =
            parts[0].trim();


        source =
            parts
                .slice(1)
                .join(".")
                .trim();
    }


    return {

        author:
            author || "Author, A.",

        year,

        title:
            title || "Article Title",

        source:
            source || "Publication Source"

    };
}


function endPeriod(str) {

    return /[.!?]$/.test(str)
        ? str
        : `${str}.`;
}


function buildStyleString(
    data,
    style,
    index = 1
) {

    const {
        author,
        year,
        title,
        source
    } = data;


    switch (style) {

        case "ieee":

            return `[${index}] ${author}, "${title}," ${source}, ${endPeriod(year)}`;


        case "apa":

            return `${author} (${year}). ${title}. ${endPeriod(source)}`;


        case "mla":

            return `${author}. "${title}." ${source}, ${endPeriod(year)}`;


        case "harvard":

            return `${author} (${year}) '${title}', ${endPeriod(source)}`;


        case "chicago":

            return `${author}. ${year}. "${title}." ${endPeriod(source)}`;


        case "vancouver":

            return `${index}. ${author}. ${title}. ${source}. ${endPeriod(year)}`;


        default:

            return `${author} (${year}). ${title}. ${endPeriod(source)}`;
    }
}
