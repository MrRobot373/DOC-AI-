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

// Point the PDF.js worker at the CDN copy bundled with pdfjs-dist.
pdfjs.GlobalWorkerOptions.workerSrc = new URL(
    "pdfjs-dist/build/pdf.worker.min.mjs",
    import.meta.url,
).toString()

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

// ── docx-preview helpers ─────────────────────────────────────────────────────

function normText(s: string): string {
    return (s || "").replace(/\s+/g, " ").trim().toLowerCase()
}

function clearDocxHighlights(container: HTMLElement) {
    container.querySelectorAll("mark[data-docai]").forEach((m) => {
        const parent = m.parentNode
        if (!parent) return
        parent.replaceChild(document.createTextNode(m.textContent || ""), m)
        parent.normalize()
    })
}

function highlightInDocx(container: HTMLElement, evidence: string) {
    clearDocxHighlights(container)
    const words = normText(evidence).split(" ").filter(Boolean).slice(0, 8)
    if (!words.length) return
    const pattern = new RegExp(words.map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("\\s+"), "i")
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT)
    let node: Node | null
    while ((node = walker.nextNode())) {
        const text = node.textContent || ""
        const match = pattern.exec(text)
        if (!match) continue
        const range = document.createRange()
        range.setStart(node, match.index)
        range.setEnd(node, match.index + match[0].length)
        const mark = document.createElement("mark")
        mark.setAttribute("data-docai", "1")
        mark.style.cssText = "background:#fde047;color:#000;padding:1px 2px;border-radius:2px"
        try {
            range.surroundContents(mark)
            mark.scrollIntoView({ behavior: "smooth", block: "center" })
        } catch {
            ;(node.parentElement as HTMLElement)?.scrollIntoView({ behavior: "smooth", block: "center" })
        }
        return
    }
}

// ── PDF text-layer highlight ─────────────────────────────────────────────────

function highlightInPdf(container: HTMLElement, evidence: string) {
    const words = normText(evidence).split(" ").filter(Boolean).slice(0, 8)
    if (!words.length) return
    const spans = Array.from(container.querySelectorAll(".react-pdf__Page__textContent span"))
    const pattern = new RegExp(words.map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("\\s+"), "i")
    for (const span of spans) {
        if (pattern.test(span.textContent || "")) {
            span.setAttribute("data-docai", "1")
            ;(span as HTMLElement).style.cssText =
                "background:rgba(253,224,71,0.5);border-radius:2px;outline:2px solid #f59e0b"
            span.scrollIntoView({ behavior: "smooth", block: "center" })
            break
        }
    }
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
    }, [file, usePdf])

    // Jump to target page in the PDF viewer.
    useEffect(() => {
        if (usePdf && targetPage && targetPage >= 1) setCurrentPage(targetPage)
    }, [targetPage, highlightNonce, usePdf])

    // Highlight text in whichever viewer is active.
    useEffect(() => {
        if (!highlight) return
        if (usePdf && pdfContainerRef.current) {
            // Small delay so the page has rendered its text layer.
            const timer = setTimeout(() => {
                if (pdfContainerRef.current) highlightInPdf(pdfContainerRef.current, highlight)
            }, 400)
            return () => clearTimeout(timer)
        } else if (!usePdf && docxContainerRef.current) {
            highlightInDocx(docxContainerRef.current, highlight)
        }
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
                {pdfLoading && (
                    <div className="flex items-center justify-center h-32 gap-2 text-gray-500 text-sm">
                        <Loader2 className="h-4 w-4 animate-spin" /> Loading PDF…
                    </div>
                )}

                {usePdf && pdfUrl ? (
                    <div ref={pdfContainerRef} className="flex justify-center py-3 bg-[#1a1a2e]">
                        <Document
                            file={pdfUrl}
                            onLoadSuccess={({ numPages: n }) => setNumPages(n)}
                            loading={
                                <div className="flex items-center gap-2 p-8 text-gray-500 text-sm">
                                    <Loader2 className="h-4 w-4 animate-spin" /> Rendering…
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
                ) : !pdfLoading ? (
                    <div
                        ref={docxContainerRef}
                        className="docx-viewer bg-white text-black p-4 text-sm"
                    />
                ) : null}
            </div>
        </div>
    )
}
