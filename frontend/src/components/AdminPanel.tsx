import { useEffect, useState } from "react"
import { supabase } from "@/lib/supabase"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { X, Trash2, Plus, Users, KeyRound, Activity } from "lucide-react"

interface AdminPanelProps {
    apiBase: string
    onClose: () => void
}

type PoolKey = {
    id: string; label?: string; provider: string; host_url: string
    model_hint?: string; vision_model_hint?: string; priority: number; active: boolean
}

export default function AdminPanel({ apiBase, onClose }: AdminPanelProps) {
    const [tab, setTab] = useState<"users" | "keys" | "usage">("keys")
    const [users, setUsers] = useState<any[]>([])
    const [keys, setKeys] = useState<PoolKey[]>([])
    const [usage, setUsage] = useState<any[]>([])
    const [loading, setLoading] = useState(false)

    // New-key form
    const [form, setForm] = useState({
        label: "", provider: "ollama_cloud", host_url: "https://ollama.com",
        api_key: "", model_hint: "", vision_model_hint: "", priority: "0",
    })

    const authHeaders = async (): Promise<Record<string, string>> => {
        const { data } = await supabase.auth.getSession()
        const t = data?.session?.access_token
        return t ? { Authorization: `Bearer ${t}` } : {}
    }

    const load = async () => {
        setLoading(true)
        try {
            const ah = await authHeaders()
            const [k, u, us] = await Promise.all([
                fetch(`${apiBase}/api/admin/pool-keys`, { headers: ah }).then(r => r.json()).catch(() => ({})),
                fetch(`${apiBase}/api/admin/users`, { headers: ah }).then(r => r.json()).catch(() => ({})),
                fetch(`${apiBase}/api/admin/usage`, { headers: ah }).then(r => r.json()).catch(() => ({})),
            ])
            setKeys(k.keys || [])
            setUsers(u.users || [])
            setUsage(us.usage || [])
        } finally {
            setLoading(false)
        }
    }

    useEffect(() => { load() }, [])

    const addKey = async () => {
        const ah = await authHeaders()
        await fetch(`${apiBase}/api/admin/pool-keys`, {
            method: "POST",
            headers: { "Content-Type": "application/json", ...ah },
            body: JSON.stringify({ ...form, priority: parseInt(form.priority) || 0 }),
        })
        setForm({ ...form, label: "", api_key: "", model_hint: "" })
        load()
    }

    const deleteKey = async (id: string) => {
        const ah = await authHeaders()
        await fetch(`${apiBase}/api/admin/pool-keys?id=${id}`, { method: "DELETE", headers: ah })
        load()
    }

    return (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 backdrop-blur-sm" onClick={onClose}>
            <div className="bg-[#0a0a0a] border border-white/10 rounded-2xl w-full max-w-3xl mx-4 shadow-2xl max-h-[85vh] overflow-hidden flex flex-col" onClick={e => e.stopPropagation()}>
                <div className="flex items-center justify-between px-6 py-4 border-b border-white/5">
                    <h2 className="text-base font-semibold text-white">Admin Panel</h2>
                    <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-white/5"><X className="h-4 w-4 text-gray-500" /></button>
                </div>

                {/* Tabs */}
                <div className="flex gap-1 px-6 pt-3">
                    {[
                        { id: "keys", label: "LLM Pool Keys", icon: KeyRound },
                        { id: "users", label: "Users", icon: Users },
                        { id: "usage", label: "Usage", icon: Activity },
                    ].map(t => (
                        <button key={t.id} onClick={() => setTab(t.id as any)}
                            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium ${tab === t.id ? "bg-white/10 text-white" : "text-gray-500 hover:text-gray-300"}`}>
                            <t.icon className="h-3.5 w-3.5" />{t.label}
                        </button>
                    ))}
                </div>

                <div className="px-6 py-4 overflow-y-auto flex-1">
                    {loading && <p className="text-xs text-gray-500">Loading…</p>}

                    {tab === "keys" && (
                        <div className="space-y-4">
                            {/* Add key form */}
                            <div className="grid grid-cols-2 gap-2 bg-white/[0.02] border border-white/5 rounded-xl p-3">
                                <Input placeholder="Label" value={form.label} onChange={e => setForm({ ...form, label: e.target.value })} className="h-9 text-xs bg-white/[0.03] border-white/10" />
                                <select value={form.provider} onChange={e => setForm({ ...form, provider: e.target.value })} className="h-9 text-xs rounded-md bg-[#0a0a0a] border border-white/10 px-2">
                                    <option value="ollama_cloud">ollama_cloud</option>
                                    <option value="freellmapi">freellmapi</option>
                                    <option value="openai_compat">openai_compat</option>
                                </select>
                                <Input placeholder="Host URL" value={form.host_url} onChange={e => setForm({ ...form, host_url: e.target.value })} className="h-9 text-xs bg-white/[0.03] border-white/10" />
                                <Input placeholder="API key" type="password" value={form.api_key} onChange={e => setForm({ ...form, api_key: e.target.value })} className="h-9 text-xs bg-white/[0.03] border-white/10" />
                                <Input placeholder="Model hint (e.g. qwen3.5:397b-cloud)" value={form.model_hint} onChange={e => setForm({ ...form, model_hint: e.target.value })} className="h-9 text-xs bg-white/[0.03] border-white/10" />
                                <Input placeholder="Priority (0 = first)" value={form.priority} onChange={e => setForm({ ...form, priority: e.target.value })} className="h-9 text-xs bg-white/[0.03] border-white/10" />
                                <Button onClick={addKey} className="col-span-2 h-9 text-xs bg-white text-black hover:bg-gray-200 gap-1.5"><Plus className="h-3.5 w-3.5" />Add pool key</Button>
                            </div>
                            {/* Key list */}
                            <div className="space-y-2">
                                {keys.length === 0 && <p className="text-xs text-gray-500">No pool keys yet. Add one above so users can pick "Auto" mode.</p>}
                                {keys.map(k => (
                                    <div key={k.id} className="flex items-center justify-between bg-white/[0.02] border border-white/5 rounded-lg px-3 py-2 text-xs">
                                        <div className="text-gray-300">
                                            <span className="font-medium text-white">{k.label || k.provider}</span>
                                            <span className="text-gray-500 ml-2">{k.host_url}</span>
                                            {k.model_hint && <span className="text-gray-500 ml-2">· {k.model_hint}</span>}
                                            <span className="text-gray-600 ml-2">· priority {k.priority}{k.active ? "" : " · inactive"}</span>
                                        </div>
                                        <button onClick={() => deleteKey(k.id)} className="p-1 rounded hover:bg-red-500/10 text-red-400"><Trash2 className="h-3.5 w-3.5" /></button>
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}

                    {tab === "users" && (
                        <div className="space-y-1.5">
                            {users.map(u => (
                                <div key={u.id} className="flex items-center justify-between bg-white/[0.02] border border-white/5 rounded-lg px-3 py-2 text-xs">
                                    <span className="text-gray-300">{u.email}</span>
                                    <span className="text-gray-500">{u.review_count} reviews</span>
                                </div>
                            ))}
                            {users.length === 0 && !loading && <p className="text-xs text-gray-500">No users found (or service-role key not configured).</p>}
                        </div>
                    )}

                    {tab === "usage" && (
                        <div className="space-y-1">
                            {usage.slice().reverse().map((u, i) => (
                                <div key={i} className="flex items-center justify-between text-xs text-gray-400 border-b border-white/5 py-1">
                                    <span>{u.user_email || "—"}</span>
                                    <span className="text-gray-500">{u.action}</span>
                                    <span className="text-gray-600">{u.created_at?.slice(0, 19).replace("T", " ")}</span>
                                </div>
                            ))}
                            {usage.length === 0 && !loading && <p className="text-xs text-gray-500">No usage logged yet.</p>}
                        </div>
                    )}
                </div>
            </div>
        </div>
    )
}
