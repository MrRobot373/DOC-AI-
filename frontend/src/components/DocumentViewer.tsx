/**
 * DocumentViewer — Viewer B (PDF.js) with Viewer A (docx-preview) fallback.
 *
 * When the backend has a rendered PDF (from Gotenberg or LibreOffice), the PDF
 * is loaded with react-pdf and the viewer jumps to the finding's page with the
 * text layer used for highlighting.
 *
 * When no PDF is available (heuristic page mode / local dev without a renderer),
 * it falls back to docx-preview — the in-browser DOCX renderer.
 */
import { useEffect, useRef, useState } from "react"
import { Document, Page, pdfjs } from "react-pdf"
import { renderAsync } from "docx-preview"
import { ChevronLeft, ChevronRight, Loader2 } from "lucide-react"

import "react-pdf/dist/Page/AnnotationLayer.css"
import "react-pdf/dist/Page/TextLayer.css"

// Point the PDF.js worker at a version-matched CDN copy. Using the runtime
// pdfjs.version avoids the fragile Vite `new URL(...import.meta.url)` resolution
// that can break in production builds.
try {
    pdfjs.GlobalWorkerOptions.workerSrc =
        `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjs.version}/pdf.worker.min.mjs`
} catch {
    /* non-fatal: the docx-preview fallback still works */
}

interface DocumentViewerProps {
    /** .docx file for the fallback docx-preview renderer. */
    file: File | null
    /** The backend review_id — used to fetch the rendered PDF. */
    reviewId: string | null
    /** Base URL of the backend API. */
    apiBase: string
    /** Auth header for the PDF fetch (Supabase JWT). */
    authToken?: string
    /** Page to jump to (1-indexed, from the finding). */
    targetPage?: number
    /** Exact text snippet to highlight in the PDF text layer. */
    highlight?: string | null
    /** Increment each time the same finding is clicked again. */
    highlightNonce?: number
}

// ── Robust cross-node text locator ──────────────────────────────────────────
// docx-preview / PDF.js fragment text across many <span> runs, often with no
// whitespace at run boundaries ("Maximum"+"Ratings" -> "MaximumRatings"). So we
// search the CONCATENATED text of all nodes and map the match back to a DOM
// Range that can span nodes — then highlight it with the CSS Custom Highlight API
// (no DOM surgery) and scroll it into view.

function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

function buildEvidenceRegex(evidence: string): RegExp | null {
    const words = (evidence || "").trim().split(/\s+/).filter(Boolean).slice(0, 12)
    if (words.length === 0) return null
    // Join words with \s* so it matches across run boundaries that have no space.
    return new RegExp(words.map(escapeRegex).join("\\s*"), "i")
}

interface NodeSpan { node: Text; start: number; end: number }

function collectText(root: HTMLElement): { raw: string; spans: NodeSpan[] } {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    let raw = ""
    const spans: NodeSpan[] = []
    let n: Node | null
    while ((n = walker.nextNode())) {
        const t = n as Text
        const txt = t.textContent || ""
        if (!txt) continue
        spans.push({ node: t, start: raw.length, end: raw.length + txt.length })
        raw += txt
    }
    return { raw, spans }
}

function posToNode(spans: NodeSpan[], pos: number): { node: Text; offset: number } | null {
    for (const s of spans) {
        if (pos >= s.start && pos <= s.end) return { node: s.node, offset: pos - s.start }
    }
    return null
}

let _highlightStyleInjected = false
function ensureHighlightStyle() {
    if (_highlightStyleInjected) return
    const style = document.createElement("style")
    style.textContent = `::highlight(docai){background:#fde047;color:#000;} mark[data-docai]{background:#fde047;color:#000;border-radius:2px;}`
    document.head.appendChild(style)
    _highlightStyleInjected = true
}

/**
 * Locate `evidence` anywhere within `root` (across nodes), highlight it, scroll
 * it into view. Tries the full phrase, then shorter prefixes for resilience.
 * Returns true if something was located.
 */
function locateAndHighlight(root: HTMLElement, evidence: string): boolean {
    ensureHighlightStyle()
    try { (CSS as any)?.highlights?.delete("docai") } catch { /* ignore */ }

    const { raw, spans } = collectText(root)
    if (!raw) return false

    // Try the full evidence, then progressively shorter leading word-sets.
    const words = (evidence || "").trim().split(/\s+/).filter(Boolean)
    for (let take = Math.min(words.length, 12); take >= 2; take -= 2) {
        const re = buildEvidenceRegex(words.slice(0, take).join(" "))
        if (!re) continue
        const m = re.exec(raw)
        if (!m) continue
        const startPos = m.index
        const endPos = m.index + m[0].length
        const a = posToNode(spans, startPos)
        const b = posToNode(spans, endPos)
        if (!a || !b) continue
        try {
            const range = document.createRange()
            range.setStart(a.node, a.offset)
            range.setEnd(b.node, b.offset)
            // CSS Custom Highlight API (cross-node, no DOM mutation).
            if ((CSS as any)?.highlights && typeof (window as any).Highlight === "function") {
                ;(CSS as any).highlights.set("docai", new (window as any).Highlight(range))
            }
            // Scroll the match into view regardless of highlight support.
            const target = (a.node.parentElement as HTMLElement) || root
            target.scrollIntoView({ behavior: "smooth", block: "center" })
            // Brief outline pulse on the containing element as a visual cue.
            target.style.transition = "background 0.2s"
            const prev = target.style.background
            target.style.background = "rgba(253,224,71,0.25)"
            setTimeout(() => { target.style.background = prev }, 1800)
            return true
        } catch {
            const target = (a.node.parentElement as HTMLElement)
            target?.scrollIntoView({ behavior: "smooth", block: "center" })
            return true
        }
    }
    return false
}

// ── Component ────────────────────────────────────────────────────────────────

export default function DocumentViewer({
    file,
    reviewId,
    apiBase,
    authToken,
    targetPage,
    highlight,
    highlightNonce,
}: DocumentViewerProps) {
    const [pdfUrl, setPdfUrl] = useState<string | null>(null)
    const [pdfLoading, setPdfLoading] = useState(false)
    const [numPages, setNumPages] = useState<number>(0)
    const [currentPage, setCurrentPage] = useState(1)
    const [usePdf, setUsePdf] = useState(false)
    const docxContainerRef = useRef<HTMLDivElement>(null)
    const pdfContainerRef = useRef<HTMLDivElement>(null)

    // Try to fetch the rendered PDF from the backend.
    useEffect(() => {
        if (!reviewId) return
        setPdfLoading(true)
        const headers: Record<string, string> = {}
        if (authToken) headers["Authorization"] = `Bearer ${authToken}`
        fetch(`${apiBase}/api/pdf/${reviewId}`, { headers })
            .then((r) => {
                if (!r.ok) throw new Error("no pdf")
                return r.blob()
            })
            .then((blob) => {
                setPdfUrl(URL.createObjectURL(blob))
                setUsePdf(true)
            })
            .catch(() => setUsePdf(false))
            .finally(() => setPdfLoading(false))
    }, [reviewId, apiBase, authToken])

    // Render docx-preview fallback when no PDF.
    useEffect(() => {
        if (usePdf || !file || !docxContainerRef.current) return
        const container = docxContainerRef.current
        container.innerHTML = ""
        file.arrayBuffer()
            .then((buf) => renderAsync(buf, container, undefined, { inWrapper: false, ignoreWidth: true }))
            .catch(() => {
                container.innerHTML = '<p style="color:#888;padding:1rem">Preview unavailable.</p>'
            })
    // Re-run when pdfLoading settles (the container is always mounted now, but
    // we need to re-render the docx whenever the PDF fetch finishes and usePdf stays false).
    }, [file, usePdf, pdfLoading])

    // Jump to target page in the PDF viewer.
    useEffect(() => {
        if (usePdf && targetPage && targetPage >= 1) setCurrentPage(targetPage)
    }, [targetPage, highlightNonce, usePdf])

    // Highlight text in whichever viewer is active. The container may still be
    // rendering (docx-preview / PDF text layer), so retry a few times.
    useEffect(() => {
        if (!highlight) return
        let cancelled = false
        let attempts = 0
        const tick = () => {
            if (cancelled) return
            const root = usePdf ? pdfContainerRef.current : docxContainerRef.current
            const found = root ? locateAndHighlight(root, highlight) : false
            attempts += 1
            if (!found && attempts < 12) setTimeout(tick, 300)  // up to ~3.6s while it renders
        }
        const start = setTimeout(tick, usePdf ? 400 : 150)
        return () => { cancelled = true; clearTimeout(start) }
    }, [highlight, highlightNonce, usePdf, currentPage])

    // ── Render ────────────────────────────────────────────────────────────────

    if (!file) return null

    return (
        <div className="rounded-xl border border-white/10 bg-[#0a0a0a] overflow-hidden flex flex-col max-h-[72vh]">
            {/* toolbar */}
            <div className="flex items-center justify-between px-3 py-2 border-b border-white/5 text-xs text-gray-500">
                <span>
                    {usePdf ? (
                        <>PDF viewer <span className="text-green-500/70 ml-1">(high-fidelity)</span></>
                    ) : (
                        <>DOCX preview <span className="text-amber-500/70 ml-1">(fast, layout approximate)</span></>
                    )}
                </span>
                {usePdf && numPages > 0 && (
                    <div className="flex items-center gap-2">
                        <button
                            onClick={() => setCurrentPage((p) => Math.max(p - 1, 1))}
                            disabled={currentPage <= 1}
                            className="p-0.5 rounded hover:bg-white/10 disabled:opacity-30"
                        >
                            <ChevronLeft className="h-3.5 w-3.5" />
                        </button>
                        <span className="tabular-nums">
                            {currentPage} / {numPages}
                        </span>
                        <button
                            onClick={() => setCurrentPage((p) => Math.min(p + 1, numPages))}
                            disabled={currentPage >= numPages}
                            className="p-0.5 rounded hover:bg-white/10 disabled:opacity-30"
                        >
                            <ChevronRight className="h-3.5 w-3.5" />
                        </button>
                    </div>
                )}
            </div>

            {/* viewer area */}
            <div className="overflow-auto flex-1">
                {/* PDF viewer — shown when a rendered PDF is available */}
                {usePdf && pdfUrl ? (
                    <div ref={pdfContainerRef} className="flex justify-center py-3 bg-[#1a1a2e]">
                        <Document
                            file={pdfUrl}
                            onLoadSuccess={({ numPages: n }) => setNumPages(n)}
                            loading={
                                <div className="flex items-center gap-2 p-8 text-gray-500 text-sm">
                                    <Loader2 className="h-4 w-4 animate-spin" /> Rendering PDF…
                                </div>
                            }
                        >
                            <Page
                                pageNumber={currentPage}
                                renderTextLayer
                                renderAnnotationLayer={false}
                                className="shadow-xl"
                                width={580}
                            />
                        </Document>
                    </div>
                ) : null}

                {/* DOCX preview — ALWAYS mounted so docx-preview never loses its container.
                    Visible only when not using the PDF viewer. Waiting for the PDF fetch
                    (pdfLoading) shows a subtle spinner overlay instead of unmounting. */}
                <div
                    ref={docxContainerRef}
                    className="docx-viewer bg-white text-black p-4 text-sm"
                    style={{ display: usePdf && pdfUrl ? "none" : "block" }}
                >
                    {pdfLoading && (
                        <div className="flex items-center justify-center h-16 gap-2 text-gray-400 text-xs">
                            <Loader2 className="h-3 w-3 animate-spin" /> Loading…
                        </div>
                    )}
                </div>
            </div>
        </div>
    )
}
