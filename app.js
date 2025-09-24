
document.addEventListener("DOMContentLoaded", () => {
    M.FormSelect.init(document.querySelectorAll("select"));
    M.Modal.init(document.querySelectorAll(".modal"));
});

// Chart Utilities
let chartInstance = null;
function destroyChart() {
    if (chartInstance) {
        chartInstance.destroy();
        chartInstance = null;
    }
}

function getContext() {
    return document.getElementById("grafica");
}

function generateColors(labels, lightness = 65) {
    return labels.map((_, i) => `hsl(${(i * 360) / labels.length}, 70%, ${lightness}%)`);
}


// Chart Renderers
// ---------------------------
function renderBarChart(labels, data, label, options = {}) {
    destroyChart();
    chartInstance = new Chart(getContext(), {
        type: "bar",
        data: {
            labels,
            datasets: [
                {
                    label, // this will be ignored in legend since legend is hidden
                    data,
                    backgroundColor: generateColors(labels),
                },
            ],
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false, // hides legend
                },
            },
            scales: {
                y: {
                    title: {
                        display: true,
                        text: "Nº Articles", // Y-axis label
                    },
                },
            },
            ...options,
        },
    });
}

function renderPieChart(labels, data) {
    destroyChart();
    chartInstance = new Chart(getContext(), {
        type: "pie",
        data: {
            labels,
            datasets: [{ data, backgroundColor: generateColors(labels) }],
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "top" } } },
    });
}

// Specific chart wrappers
function renderChartByYear(byYear) {
    renderBarChart(Object.keys(byYear), Object.values(byYear), "Publications per Year");
}

function renderChartByType(byType) {
    renderPieChart(Object.keys(byType), Object.values(byType));
}

function renderStackedChart(years, categories, dataByYearCategory) {
    destroyChart();
    const fixedOrder = ["Journal", "Conference", "arXiv"];
    const colorMap = {
        Journal: "hsl(120, 70%, 80%)",
        Conference: "hsl(220, 70%, 80%)",
        arXiv: "hsl(0, 70%, 80%)",
    };
    const datasets = fixedOrder.map((cat) => ({
        label: cat,
        data: years.map((year) => dataByYearCategory[year]?.[cat] || 0),
        backgroundColor: colorMap[cat] || "hsl(0, 0%, 80%)",
    }));
    chartInstance = new Chart(getContext(), {
        type: "bar",
        data: { labels: years, datasets },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: "top" } },
            scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true } },
        },
    });
}

function renderHorizontalStackedChart(labels, datasetMap, order, colorMap, xTitle, yTitle) {
    destroyChart();
    const datasets = order.map((key) => ({
        label: key,
        data: labels.map((lbl) => datasetMap[lbl]?.[key] || 0),
        backgroundColor: colorMap[key] || "hsl(0, 0%, 80%)",
    }));
    chartInstance = new Chart(getContext(), {
        type: "bar",
        data: { labels, datasets },
        options: {
            indexAxis: "y",
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: "top" } },
            scales: {
                x: { stacked: true, beginAtZero: true, title: { display: true, text: xTitle } },
                y: { stacked: true, title: { display: true, text: yTitle } },
            },
        },
    });
}


// Abstract Modal
function showAbstract(title, abstract) {
    document.getElementById("modal-abstract-title").textContent = title;
    document.getElementById("modal-abstract-content").textContent = abstract;
    M.Modal.getInstance(document.getElementById("modal-abstract")).open();
}

function showAbstractFromAttr(el) {
    showAbstract(
        decodeURIComponent(el.getAttribute("data-title") || ""),
        decodeURIComponent(el.getAttribute("data-abstract") || "")
    );
}


// CSV / JSON / XLSX Export
let globalDataArray = null;
let globalHeaders = null;

fetch("data/Papers.csv")
    .then((res) => res.text())
    .then((text) => {
        document.getElementById("download-csv").href =
            "data:text/csv;charset=utf-8," + encodeURIComponent(text);
    });

Papa.parse("data/Papers.csv", {
    download: true,
    header: true,
    skipEmptyLines: true,
    complete: ({ data, meta }) => {
        const headers = meta.fields;
        data = data.filter((row) => row.ID?.trim());

        globalDataArray = data;
        globalHeaders = headers;

        // JSON
        const jsonUrl = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: "application/json" }));
        document.getElementById("download-json").href = jsonUrl;

        // XLSX
        document.getElementById("download-xlsx").addEventListener("click", (e) => {
            e.preventDefault();
            if (!window.XLSX) {
                const script = document.createElement("script");
                script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
                script.onload = () => exportToXLSX(globalDataArray, globalHeaders);
                document.body.appendChild(script);
            } else {
                exportToXLSX(globalDataArray, globalHeaders);
            }
        });

        initDataTable(data, headers);
        processCharts(data);
    },
});

function exportToXLSX(dataArray, headers) {
    const ws = XLSX.utils.json_to_sheet(dataArray, { header: headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Papers");
    XLSX.writeFile(wb, "Papers.xlsx");
}


// DataTable
function initDataTable(data, headers) {
    const dt = $("#tabla").DataTable({
        data: data.map((row) => headers.map((h) => row[h] || "")),
        columns: headers.map((h) => ({ title: h })),
        pageLength: 5,
        scrollX: true,
        columnDefs: [
            {
                targets: headers.indexOf("KEY"),
                visible: false,
                searchable: false,
            },
            {
                targets: headers.indexOf("TITLE"),
                width: "30%",
                render: (data, _, row) => {
                    const bibtex = row[headers.indexOf("BIBTEX")] || "";
                    const url = bibtex.match(/url\s*=\s*[{"]([^}"]+)[}"]/i)?.[1];
                    return url ? `<a href="${url}" target="_blank">${data}</a>` : data;
                },
            },
            {
                targets: headers.indexOf("BIBTEX"),
                render: (data, _, row) => {
                    if (!data) return "";
                    const fileName = `${row[headers.indexOf("KEY")] || "citation"}.bib`;
                    const url = URL.createObjectURL(new Blob([data], { type: "text/plain" }));
                    return `<a class="green-btn" href="${url}" download="${fileName}">DOWNLOAD</a>`;
                },
            },
            {
                targets: headers.indexOf("ABSTRACT"),
                render: (data, _, row) =>
                    data
                        ? `<a class="green-btn" href="#modal-abstract"
                            data-title="${encodeURIComponent(row[headers.indexOf("TITLE")] || "")}"
                            data-abstract="${encodeURIComponent(data)}"
                            onclick="showAbstractFromAttr(this)">
                            INFO</a>`
                        : "",
            },
        ],
    });

    M.FormSelect.init(document.querySelectorAll("#pageLengthSelect"));
    document.getElementById("pageLengthSelect").addEventListener("change", (e) => {
        dt.page.len(Number(e.target.value)).draw();
    });
}

// Data Processing for Charts
function processCharts(data) {
    const byYear = {};
    const byType = {};
    const llmsUsedCount = {};
    const benchmarksCount = {};
    const metricsCount = {};
    const classificationLLMType = {};
    const llmApproachTypesSet = new Set();
    const conferencesCount = {};
    const journalsCount = {};

    data.forEach((r) => {
        // Year & type
        if (r.YEAR) byYear[r.YEAR] = (byYear[r.YEAR] || 0) + 1;
        if (r["PUBLICATION TYPE"]) byType[r["PUBLICATION TYPE"]] = (byType[r["PUBLICATION TYPE"]] || 0) + 1;

        // LLMs
        if (r["LLMs USED"]) {
            r["LLMs USED"].split(",").map((s) => s.trim()).forEach((m) => {
                if (m && !["n/s", "none"].includes(m.toLowerCase())) {
                    llmsUsedCount[m] = (llmsUsedCount[m] || 0) + 1;
                }
            });
        }

        // Benchmarks
        if (r.BENCHMARK && !["none", "no bmk-ds"].includes(r.BENCHMARK.toLowerCase())) {
            r.BENCHMARK.split(",").map((s) => s.trim()).forEach((b) => {
                if (b) benchmarksCount[b] = (benchmarksCount[b] || 0) + 1;
            });
        }

        // Metrics
        if (r["EVALUATION METRIC"] && !["none", "no eval."].includes(r["EVALUATION METRIC"].toLowerCase())) {
            r["EVALUATION METRIC"].split(",").map((s) => s.trim()).forEach((m) => {
                if (m) metricsCount[m] = (metricsCount[m] || 0) + 1;
            });
        }

        // Classification vs LLM Approach
        if (r.CATEGORY && r["LLM APPROACH TYPE"]) {
            r.CATEGORY.split(",").map((s) => s.trim()).forEach((cl) => {
                if (!classificationLLMType[cl]) classificationLLMType[cl] = {};
                r["LLM APPROACH TYPE"].split(",").map((s) => s.trim()).forEach((t) => {
                    classificationLLMType[cl][t] = (classificationLLMType[cl][t] || 0) + 1;
                    llmApproachTypesSet.add(t);
                });
            });
        }

        // Conferences / Journals
        if (r["PUBLISHED INTO"]) {
            const val = r["PUBLISHED INTO"].trim();
            if (val.startsWith("C:")) conferencesCount[val.replace(/^C:\s*/, "")] = (conferencesCount[val.replace(/^C:\s*/, "")] || 0) + 1;
            if (val.startsWith("J:")) journalsCount[val.replace(/^J:\s*/, "")] = (journalsCount[val.replace(/^J:\s*/, "")] || 0) + 1;
        }
    });

    // Year/category stacked data
    const years = [...new Set(data.map((r) => r.YEAR).filter(Boolean))].sort();
    const categories = [...new Set(data.map((r) => r["PUBLICATION TYPE"]).filter(Boolean))];
    const dataByYearCategory = {};
    data.forEach((r) => {
        if (r.YEAR && r["PUBLICATION TYPE"]) {
            dataByYearCategory[r.YEAR] ??= {};
            dataByYearCategory[r.YEAR][r["PUBLICATION TYPE"]] =
                (dataByYearCategory[r.YEAR][r["PUBLICATION TYPE"]] || 0) + 1;
        }
    });

    // Initial render
    renderStackedChart(years, categories, dataByYearCategory);

    // Chart selector
    document.getElementById("graficoSelect").addEventListener("change", (e) => {
        switch (e.target.value) {
            case "ano": renderStackedChart(years, categories, dataByYearCategory); break;
            case "llmsused": renderBarChart(Object.keys(llmsUsedCount), Object.values(llmsUsedCount), "LLMs Used"); break;
            case "benchmarks": renderBarChart(Object.keys(benchmarksCount), Object.values(benchmarksCount), "Benchmarks"); break;
            case "metrics": renderBarChart(Object.keys(metricsCount), Object.values(metricsCount), "Metrics"); break;
            case "category":
                const order = ["LLM-Pure-Prompting", "Hybrid-Prompting", "LLM-Pure-FineTune", "Hybrid-FineTune", "None"];
                const colorMap = {
                    "LLM-Pure-Prompting": "hsl(200,70%,80%)",
                    "Hybrid-Prompting": "hsl(200,70%,40%)",
                    "LLM-Pure-FineTune": "hsl(0,70%,80%)",
                    "Hybrid-FineTune": "hsl(0,70%,40%)",
                    "None": "hsl(0,0%,80%)",
                };
                const sortedCats = Object.entries(classificationLLMType)
                    .map(([cl, vals]) => [cl, Object.values(vals).reduce((a, b) => a + b, 0)])
                    .sort((a, b) => b[1] - a[1])
                    .map(([cl]) => cl);
                renderHorizontalStackedChart(sortedCats, classificationLLMType, order, colorMap, "Nº Articles", "Categories");
                break;
            case "conferences": renderPieChart(Object.keys(conferencesCount), Object.values(conferencesCount)); break;
            case "journals": renderPieChart(Object.keys(journalsCount), Object.values(journalsCount)); break;
        }
    });
}

// Ensure modal and BibTeX buttons work
document.addEventListener("DOMContentLoaded", function () {
    // Initialize modal
    const citationModalElem = document.getElementById("modal-citation");
    const citationModal = M.Modal.init(citationModalElem);

    const bibtexBtn = document.getElementById("download-bibtex");
    const showCitationBtn = document.getElementById("show-citation");
    const copyCitationBtn = document.getElementById("copy-citation");
    const citationContent = document.getElementById("citation-content");

    // Fetch CITATION.cff once
    let citationData = null;
    fetch("CITATION.cff")
        .then((res) => res.text())
        .then((text) => {
            const cff = jsyaml.load(text);
            const pc = cff["preferred-citation"];
            citationData = {
                authors: pc.authors.map((a) => ({ given: a["given-names"], family: a["family-names"], orcid: a.orcid || null })),
                title: pc.title,
                journal: pc.journal?.name || "",
                year: pc.year,
                volume: pc.journal?.volume || "",
                issue: pc.journal?.issue || "",
                pages: pc.journal?.pages ? `${pc.journal.pages.start}-${pc.journal.pages.end}` : "",
                publisher: pc.journal?.publisher || "",
                doi: pc.doi || "",
                url: cff.url || "",
            };
        })
        .catch((err) => console.error("Error loading CITATION.cff:", err));

    // Helper to generate BibTeX
    function formatBibTeX(data) {
        if (!data) return "";
        const authors = data.authors.map(a => `${a.family}, ${a.given}`).join(" and ");
        return `@article{Augusto2026,
                author    = {${authors}},
                title     = {${data.title}},
                journal   = {${data.journal}},
                year      = {${data.year}},
                volume    = {${data.volume}},
                number    = {${data.issue}},
                pages     = {${data.pages}},
                publisher = {${data.publisher}},
                doi       = {${data.doi}}
            }`;
    }

    function formatCitationAPA(data) {
        if (!data) return "";
        const authors = data.authors
            .map(a => `${a.family}, ${a.given[0]}.`)
            .join(", ")
            .replace(/, ([^,]*)$/, ", & $1");
        return `${authors} (${data.year}). ${data.title}. <i>${data.journal}</i>, ${data.volume}(${data.issue}), ${data.pages}. ${data.publisher}. https://doi.org/${data.doi}`;
    }

    // BibTeX download
bibtexBtn.addEventListener("click", function (e) {
    e.preventDefault();
    if (!citationData) return alert("Citation not loaded yet.");

    const bibtexStr = formatBibTeX(citationData);
    const blob = new Blob([bibtexStr], { type: "text/x-bibtex" });
    const url = URL.createObjectURL(blob);

    // Create temporary link and trigger download
    const tmpLink = document.createElement("a");
    tmpLink.href = url;
    tmpLink.download = "Augusto2026.bib";
    document.body.appendChild(tmpLink);
    tmpLink.click();
    document.body.removeChild(tmpLink);

    // Revoke object URL after a short delay to ensure download starts
    setTimeout(() => URL.revokeObjectURL(url), 1000);
});
    // Show modal with APA citation
    showCitationBtn.addEventListener("click", function () {
        if (!citationData) return alert("Citation not loaded yet.");
        citationContent.innerHTML = formatCitationAPA(citationData);
        citationModal.open();
    });

    // Copy citation to clipboard
    copyCitationBtn.addEventListener("click", function (e) {
        e.preventDefault();
        if (!citationData) return alert("Citation not loaded yet.");
        const text = citationContent.innerText || "";
        navigator.clipboard.writeText(text).then(() => M.toast({ html: "Citation copied!" }));
    });
});
