import { useEffect, useState } from "react"
import { supabase } from "@/lib/supabase"
import { Button } from "@/components/ui/button"
import { X, Download, GitCompare, FileText } from "lucide-react"

interface HistoryPanelProps {
    apiBase: string
    userId: string
    onClose: () => void
}

type HistoryRow = {
    id: string
    document_name: string
    report_filename: string
    review_id?: string
    created_at: string
}

export default function HistoryPanel({ apiBase, userId, onClose }: HistoryPanelProps) {
    const [rows, setRows] = useState<HistoryRow[]>([])
    const [loading, setLoading] = useState(false)
    const [compareSel, setCompareSel] = useState<string[]>([])  // up to 2 review_ids
    const [diff, setDiff] = useState<any | null>(null)
    const [diffLoading, setDiffLoading] = useState(false)

    const authHeaders = async (): Promise<Record<string, string>> => {
        const { data } = await supabase.auth.getSession()
        const t = data?.session?.access_token
        return t ? { Authorization: `Bearer ${t}` } : {}
    }

    useEffect(() => {
        let cancelled = false
        ;(async () => {
            setLoading(true)
            try {
                const { data } = await supabase.from("review_history").select("*")
                    .eq("user_id", userId).order("created_at", { ascending: false }).limit(30)
                if (!cancelled) setRows(data || [])
            } catch {
                /* ignore */
            } finally {
                if (!cancelled) setLoading(false)
            }
        })()
        return () => { cancelled = true }
    }, [userId])

    const toggleCompare = (reviewId?: string) => {
        if (!reviewId) return
        setCompareSel(prev => {
            if (prev.includes(reviewId)) return prev.filter(x => x !== reviewId)
            if (prev.length >= 2) return [prev[1], reviewId]  // keep last 2
            return [...prev, reviewId]
        })
    }

    const runCompare = async () => {
        if (compareSel.length !== 2) return
        setDiffLoading(true)
        setDiff(null)
        try {
            // compareSel[1] is the most-recently selected = newer; [0] = older.
            const [a, b] = compareSel
            const ah = await authHeaders()
            const resp = await fetch(`${apiBase}/api/compare?old=${a}&new=${b}`, { headers: ah })
            setDiff(await resp.json())
        } finally {
            setDiffLoading(false)
        }
    }

    return (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 backdrop-blur-sm" onClick={onClose}>
            <div className="bg-[#0a0a0a] border border-white/10 rounded-2xl w-full max-w-2xl mx-4 shadow-2xl max-h-[85vh] overflow-hidden flex flex-col" onClick={e => e.stopPropagation()}>
                <div className="flex items-center justify-between px-6 py-4 border-b border-white/5">
                    <h2 className="text-base font-semibold text-white">Review History</h2>
                    <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-white/5"><X className="h-4 w-4 text-gray-500" /></button>
                </div>

                <div className="px-6 py-3 border-b border-white/5 flex items-center justify-between">
                    <p className="text-xs text-gray-500">Select 2 reviews to compare versions.</p>
                    <Button onClick={runCompare} disabled={compareSel.length !== 2 || diffLoading}
                        className="h-8 text-xs gap-1.5 bg-blue-600/20 text-blue-400 hover:bg-blue-600/30 border border-blue-500/20" variant="ghost">
                        <GitCompare className="h-3.5 w-3.5" />{diffLoading ? "Comparing…" : "Compare selected"}
                    </Button>
                </div>

                <div className="px-6 py-4 overflow-y-auto flex-1 space-y-2">
                    {loading && <p className="text-xs text-gray-500">Loading…</p>}
                    {!loading && rows.length === 0 && <p className="text-xs text-gray-500">No past reviews yet.</p>}

                    {rows.map(r => (
                        <div key={r.id} className="flex items-center justify-between bg-white/[0.02] border border-white/5 rounded-lg px-3 py-2 text-xs">
                            <label className="flex items-center gap-2 cursor-pointer min-w-0">
                                <input type="checkbox" checked={!!r.review_id && compareSel.includes(r.review_id)}
                                    onChange={() => toggleCompare(r.review_id)} disabled={!r.review_id}
                                    className="h-3.5 w-3.5 accent-blue-500" />
                                <FileText className="h-3.5 w-3.5 text-gray-500 shrink-0" />
                                <span className="text-gray-300 truncate">{r.document_name}</span>
                                <span className="text-gray-600 shrink-0">{r.created_at?.slice(0, 10)}</span>
                            </label>
                            <a href={`${apiBase}/api/download/${r.report_filename}`} download
                                className="flex items-center gap-1 text-gray-400 hover:text-white shrink-0">
                                <Download className="h-3.5 w-3.5" />
                            </a>
                        </div>
                    ))}

                    {diff && (
                        <div className="mt-4 border-t border-white/10 pt-3">
                            <p className="text-xs font-medium text-white mb-2">
                                Comparison: <span className="text-green-400">{diff.summary?.new ?? 0} new</span> ·{" "}
                                <span className="text-gray-400">{diff.summary?.fixed ?? 0} fixed</span> ·{" "}
                                <span className="text-gray-500">{diff.summary?.unchanged ?? 0} unchanged</span>
                            </p>
                            {(diff.new || []).slice(0, 20).map((f: any, i: number) => (
                                <div key={i} className="text-xs text-gray-300 border-l-2 border-green-500/40 pl-2 py-1 mb-1">
                                    <span className="text-green-400 font-medium">NEW</span> · {f.category?.replace(/_/g, " ")}: {f.comment?.slice(0, 120)}
                                </div>
                            ))}
                        </div>
                    )}
                </div>
            </div>
        </div>
    )
}
