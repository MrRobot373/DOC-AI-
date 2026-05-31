import { useEffect, useRef } from "react"
import { renderAsync } from "docx-preview"

interface DocumentViewerProps {
    /** The uploaded .docx file to render (Word only). */
    file: File | null
    /** Exact text to locate and highlight (a finding's evidence quote). */
    highlight: string | null
    /** Monotonic counter so re-clicking the same finding re-triggers a scroll. */
    highlightNonce?: number
}

function normalize(s: string): string {
    return s.replace(/\s+/g, " ").trim().toLowerCase()
}

function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

function clearHighlights(container: HTMLElement) {
    container.querySelectorAll("mark[data-docai]").forEach((m) => {
        const parent = m.parentNode
        if (!parent) return
        parent.replaceChild(document.createTextNode(m.textContent || ""), m)
        parent.normalize()
    })
}

/**
 * Find the first text node whose content matches the (whitespace-tolerant)
 * evidence snippet, wrap the match in <mark>, and scroll it into view.
 */
function highlightEvidence(container: HTMLElement, evidence: string) {
    clearHighlights(container)
    const words = normalize(evidence).split(" ").filter(Boolean).slice(0, 10)
    if (words.length === 0) return
    // Match the leading words allowing any whitespace between them.
    const pattern = new RegExp(words.map(escapeRegex).join("\\s+"), "i")

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
        mark.style.backgroundColor = "#fde047"
        mark.style.color = "#000"
        mark.style.padding = "1px 2px"
        mark.style.borderRadius = "2px"
        try {
            range.surroundContents(mark)
            mark.scrollIntoView({ behavior: "smooth", block: "center" })
        } catch {
            // surroundContents fails if the range crosses element boundaries;
            // fall back to scrolling the containing element into view.
            const el = (node.parentElement as HTMLElement) || null
            el?.scrollIntoView({ behavior: "smooth", block: "center" })
        }
        return
    }
}

export default function DocumentViewer({ file, highlight, highlightNonce }: DocumentViewerProps) {
    const containerRef = useRef<HTMLDivElement>(null)

    // Render the document whenever the file changes.
    useEffect(() => {
        const container = containerRef.current
        if (!container) return
        container.innerHTML = ""
        if (!file) return
        file
            .arrayBuffer()
            .then((buf) => renderAsync(buf, container, undefined, { inWrapper: false, ignoreWidth: true }))
            .catch(() => {
                container.innerHTML =
                    '<p style="color:#888;padding:1rem">Preview unavailable for this file.</p>'
            })
    }, [file])

    // Highlight + scroll whenever the requested evidence changes.
    useEffect(() => {
        const container = containerRef.current
        if (!container || !highlight) return
        highlightEvidence(container, highlight)
    }, [highlight, highlightNonce])

    return (
        <div className="docx-viewer rounded-xl border border-white/10 bg-white text-black overflow-auto max-h-[70vh] p-4">
            <div ref={containerRef} />
        </div>
    )
}
