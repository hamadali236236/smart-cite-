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


/**
 * Normalizes characters that Word's AutoCorrect/AutoFormat commonly
 * substitutes (smart quotes, en/em dashes, non-breaking spaces) so the
 * regex-based parsers downstream see predictable, literal characters.
 * Run this on every piece of text pulled out of the document BEFORE
 * it hits parseCitationText / KNOWN_STYLE_PATTERNS.
 */
function sanitizeText(text) {
    return text
        .replace(/[\u2018\u2019\u201B]/g, "'")   // smart single quotes -> '
        .replace(/[\u201C\u201D\u201F]/g, '"')   // smart double quotes -> "
        .replace(/[\u2013\u2014]/g, "-")          // en/em dash -> hyphen
        .replace(/\u00A0/g, " ")                  // non-breaking space -> space
        .replace(/\s+/g, " ")
        .trim();
}


/**
 * A line is treated as the START of a new reference entry if it looks
 * like a numbered marker ("[1]", "1.", "1)") or begins with an
 * "Lastname, F." author pattern. Any other non-blank line is assumed
 * to be a wrapped continuation of the reference currently being built.
 * A blank line always closes out the current reference.
 */
const REFERENCE_START_PATTERN =
    /^(?:\[\d+\]|\d+[.)])\s*\S|^[A-Z][a-zA-Z'-]+,\s*[A-Z]\.(?:\s*[A-Z]\.)?/;


/**
 * Groups the raw selected text into one string per reference entry,
 * instead of naively treating every line break as a new entry. This
 * correctly re-joins references that wrap across multiple lines
 * (very common when pasted/typed with hanging-indent formatting).
 */
function splitIntoReferenceBlocks(text) {

    const normalized = text.replace(/[\r\v]+/g, "\n");

    const lines = normalized.split("\n");

    const blocks = [];
    let current = "";

    for (const rawLine of lines) {

        const trimmed = rawLine.trim();

        if (trimmed.length === 0) {
            if (current) {
                blocks.push(current.trim());
                current = "";
            }
            continue;
        }

        const startsNewReference =
            current.length === 0 ||
            REFERENCE_START_PATTERN.test(trimmed);

        if (startsNewReference && current.length > 0) {
            blocks.push(current.trim());
            current = trimmed;
        } else {
            current = current
                ? `${current} ${trimmed}`
                : trimmed;
        }
    }

    if (current) {
        blocks.push(current.trim());
    }

    return blocks.filter((b) => b.length > 0);
}


async function formatSelection() {
    setLoading(true);

    try {
        await Word.run(async (context) => {

            const selection = context.document.getSelection();

            selection.load("text");

            await context.sync();


            const rawLines = splitIntoReferenceBlocks(selection.text)
                .map((block) => sanitizeText(block))
                .filter((block) => block.length > 0);


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
 * Pulls every "LastName, X." author out of a parsed author string,
 * in original order. Falls back to a single guessed last name if the
 * string doesn't match the usual "Last, F." pattern (e.g. already a
 * bare name, or an unusual format).
 */
function extractLastNames(authorStr) {

    const matches =
        [...authorStr.matchAll(/([A-Z][a-zA-Z'-]+),\s*[A-Z]\.(?:\s*[A-Z]\.)?/g)];

    if (matches.length > 0) {
        return matches.map((m) => m[1]);
    }

    const fallback = authorStr.split(",")[0].trim();

    return fallback ? [fallback] : ["Author"];
}


/**
 * Builds the "best" in-text author form used when WRITING a new
 * citation marker (Chicago/MLA/APA/Harvard): "Smith" / "Johnson & Lee"
 * / "Brown et al." — the conventional short form for 3+ authors.
 */
function inTextAuthorForm(authorStr) {

    const names = extractLastNames(authorStr);

    if (names.length === 1) {
        return names[0];
    }

    if (names.length === 2) {
        return `${names[0]} & ${names[1]}`;
    }

    return `${names[0]} et al.`;
}


/**
 * Builds every plausible in-text author form used when SEARCHING the
 * document for an existing citation to replace. Some reference styles
 * (e.g. APA 6th) spell out all authors up to five on first mention
 * instead of using "et al." right away, so for 3+ authors this returns
 * BOTH the short "et al." form and the full spelled-out list, so
 * either convention found in the document text gets matched.
 */
function inTextAuthorForms(authorStr) {

    const names = extractLastNames(authorStr);

    if (names.length === 1) {
        return [names[0]];
    }

    if (names.length === 2) {
        return [`${names[0]} & ${names[1]}`];
    }

    const fullList =
        `${names.slice(0, -1).join(", ")}, & ${names[names.length - 1]}`;

    const etAl = `${names[0]} et al.`;

    return [etAl, fullList];
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


    const authorForms =
        inTextAuthorForms(parsedData.author);


    const newMarker =
        buildInlineMarker(
            parsedData,
            selectedStyle,
            index
        );


    const candidates = new Set([

        `[${index}]`,

        `(${index})`

    ]);


    for (const authorForm of authorForms) {

        candidates.add(`(${authorForm}, ${parsedData.year})`);
        candidates.add(`(${authorForm} ${parsedData.year})`);
        candidates.add(`(${authorForm})`);
    }


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

    const authorForm =
        inTextAuthorForm(parsedData.author);


    switch (style) {

        case "ieee":

            return `[${index}]`;


        case "vancouver":

            return `(${index})`;


        case "chicago":

            return `(${authorForm} ${parsedData.year})`;


        case "mla":

            return `(${authorForm})`;


        case "apa":

        case "harvard":

        default:

            return `(${authorForm}, ${parsedData.year})`;
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

    const clean = sanitizeText(text);

    for (
        const { regex, map }
        of KNOWN_STYLE_PATTERNS
    ) {

        const match =
            clean.match(regex);


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


    return parseRawCitationText(clean);
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


    // Matches one or more "Lastname, F." / "Lastname, F. M." groups,
    // separated by commas, "&", or "and" (e.g. "Brown, T., Wilson, A., & Khan, S.")
    // Matches one or more "Lastname, F." author groups, joined by any of:
    // ", & " (Oxford-comma + ampersand before the last author), a bare
    // "," or "&", or the word "and". The ",\s*&\s*" branch MUST come
    // before the plain ",\s*" branch, or the regex consumes only the
    // comma and stops at "&" (which doesn't match the next author's
    // required leading capital letter), silently dropping every author
    // after the first "&".
    const authorMatch =
        cleanText.match(
            /^([A-Z][a-zA-Z'-]+,\s*[A-Z]\.(?:\s*[A-Z]\.)?\s*(?:,\s*&\s*|,\s*|&\s*|\s+and\s+)?)+/
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
