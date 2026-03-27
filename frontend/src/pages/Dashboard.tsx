import { useState, useEffect, useRef } from "react"
import { useNavigate } from "react-router-dom"
import { supabase } from "@/lib/supabase"
import type { User } from "@supabase/supabase-js"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Card, CardContent } from "@/components/ui/card"
import {
    LogOut, Settings, UploadCloud, CheckCircle2, AlertTriangle,
    FileText, X, ChevronDown, ExternalLink, MessageSquare, HelpCircle, Send
} from "lucide-react"

interface DashboardProps {
    user: User
}

export default function Dashboard({ user }: DashboardProps) {
    const navigate = useNavigate()
    const API_BASE_URL = import.meta.env.VITE_API_URL || "http://localhost:5000"

    // Profile dropdown
    const [showProfileMenu, setShowProfileMenu] = useState(false)
    const profileRef = useRef<HTMLDivElement>(null)

    // Settings modal
    const [showSettingsModal, setShowSettingsModal] = useState(false)
    const [showFeedbackModal, setShowFeedbackModal] = useState(false)
    const [showHelpModal, setShowHelpModal] = useState(false)
    const [feedbackText, setFeedbackText] = useState("")
    const [feedbackType, setFeedbackType] = useState("bug")
    const [sendingFeedback, setSendingFeedback] = useState(false)
    const [feedbackSent, setFeedbackSent] = useState(false)
    const [feedbackImage, setFeedbackImage] = useState<File | null>(null)
    const feedbackImageRef = useRef<HTMLInputElement>(null)

    // API Settings State
    const [apiKey, setApiKey] = useState("")
    const [hostUrl, setHostUrl] = useState("https://ollama.com")
    const [savingSettings, setSavingSettings] = useState(false)
    const [testResult, setTestResult] = useState<string | null>(null)
    const [availableModels] = useState<string[]>([
        "qwen3-vl:235b-cloud",
        "kimi-k2-thinking:cloud",
        "kimi-k2.5:cloud",
        "nemotron-3-super:cloud"
    ])
    const [selectedModel, setSelectedModel] = useState("kimi-k2-thinking:cloud")
    const [visionModel, setVisionModel] = useState("qwen3-vl:235b-cloud")
    const [reviewMode, setReviewMode] = useState<"normal" | "pro">("pro")

    // Review State
    const [selectedFile, setSelectedFile] = useState<File | null>(null)
    const [reviewing, setReviewing] = useState(false)
    const [progressMsg, setProgressMsg] = useState("")
    const [progressPct, setProgressPct] = useState(0)
    const [reportUrl, setReportUrl] = useState<string | null>(null)
    const [findings, setFindings] = useState<any[]>([])
    const [fileType, setFileType] = useState<'doc' | 'excel'>('doc')

    const fileInputRef = useRef<HTMLInputElement>(null)

    // Close profile dropdown on outside click
    useEffect(() => {
        const handler = (e: MouseEvent) => {
            if (profileRef.current && !profileRef.current.contains(e.target as Node)) {
                setShowProfileMenu(false)
            }
        }
        document.addEventListener("mousedown", handler)
        return () => document.removeEventListener("mousedown", handler)
    }, [])

    // Load settings on mount
    useEffect(() => {
        const loadSettings = async () => {
            const { data } = await supabase
                .from('user_settings')
                .select('*')
                .eq('user_id', user.id)
                .single()
            if (data) {
                setApiKey(data.ollama_api_key || "")
                setHostUrl(data.ollama_host_url || "https://ollama.com")
            }
        }
        loadSettings().catch(() => { })
    }, [user.id])

    const handleLogout = async () => {
        await supabase.auth.signOut()
        navigate("/")
    }

    const handleSaveSettings = async () => {
        setSavingSettings(true)
        try {
            await supabase.from('user_settings').upsert({
                user_id: user.id,
                ollama_api_key: apiKey,
                ollama_host_url: hostUrl
            })

            const resp = await fetch(`${API_BASE_URL}/api/check-ollama`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ api_key: apiKey, host: hostUrl })
            })
            const data = await resp.json()

            if (data.success) {
                setTestResult("Connection Successful!")
            } else {
                setTestResult("Error: " + data.error)
            }
        } catch (e: any) {
            setTestResult("Error: " + e.message)
        } finally {
            setSavingSettings(false)
        }
    }

    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0]
        if (!file) return

        const ext = file.name.split('.').pop()?.toLowerCase()
        if (fileType === 'excel') {
            if (ext !== 'xlsx' && ext !== 'xls') {
                alert("Please upload an Excel file (.xlsx or .xls)")
                return
            }
        } else {
            if (ext !== 'docx' && ext !== 'doc') {
                alert("Please upload a Word document (.docx or .doc)")
                return
            }
        }

        setSelectedFile(file)
        setFindings([])
        setProgressPct(0)
        setProgressMsg("")
        setReportUrl(null)
    }

    const handleStartReview = async () => {
        if (!selectedFile || !apiKey) {
            setShowSettingsModal(true)
            return
        }

        setReviewing(true)
        setProgressMsg("Starting review process...")
        setProgressPct(5)
        setReportUrl(null)
        setFindings([])

        const formData = new FormData()
        if (user) formData.append('user_id', user.id)
        formData.append('api_key', apiKey)
        formData.append('host', hostUrl)
        formData.append('model', selectedModel || "gpt-oss:120b-cloud")
        formData.append('vision_model', visionModel)
        formData.append('review_mode', reviewMode)
        formData.append('document', selectedFile)
        formData.append('file_type', fileType)

        try {
            const resp = await fetch(`${API_BASE_URL}/api/review`, {
                method: 'POST',
                body: formData,
            })
            const data = await resp.json()

            if (!data.success) {
                setProgressMsg(`Failed: ${data.error}`)
                setReviewing(false)
                return
            }

            const reviewId = data.review_id

            const poll = setInterval(async () => {
                try {
                    const sResp = await fetch(`${API_BASE_URL}/api/progress/${reviewId}`)
                    const sData = await sResp.json()

                    if (sData.status === 'error') {
                        clearInterval(poll)
                        setProgressMsg(`Error: ${sData.message}`)
                        setReviewing(false)
                    } else if (sData.status === 'done') {
                        clearInterval(poll)
                        setProgressMsg("Review Complete!")
                        setProgressPct(100)
                        setFindings(sData.findings || [])
                        setReportUrl(`${API_BASE_URL}/api/download/${sData.report_filename}`)
                        setReviewing(false)

                        supabase.from('review_history').insert({
                            user_id: user.id,
                            document_name: selectedFile.name,
                            report_filename: sData.report_filename
                        }).then()
                    } else {
                        setProgressMsg(sData.message || "Processing...")
                        setProgressPct(sData.progress || 5)
                    }
                } catch (e) {
                    console.error(e)
                }
            }, 2000)
        } catch (e: any) {
            setProgressMsg(`Error: ${e.message}`)
            setReviewing(false)
        }
    }

    // Get user initials
    const initials = user.email
        ? user.email.substring(0, 2).toUpperCase()
        : "U"

    return (
        <div className="min-h-screen bg-black text-gray-100 flex flex-col font-sans">
            {/* ─── Header ─── */}
            <header className="border-b border-white/10 bg-black/80 backdrop-blur-md sticky top-0 z-50">
                <div className="max-w-6xl mx-auto px-6 h-16 flex items-center justify-between">
                    <div className="flex items-center gap-3">
                        <h1 className="text-xl font-bold tracking-tight text-white">DOC-AI</h1>
                        <div className="h-4 w-[1px] bg-white/20"></div>
                        <span className="font-semibold text-red-500 tracking-wider text-sm">GMS</span>
                    </div>

                    {/* Profile Button */}
                    <div className="relative" ref={profileRef}>
                        <button
                            onClick={() => setShowProfileMenu(!showProfileMenu)}
                            className="flex items-center gap-2 px-3 py-1.5 rounded-full border border-white/10 hover:bg-white/5 transition-colors"
                        >
                            <div className="h-7 w-7 rounded-full bg-white/10 flex items-center justify-center text-xs font-medium text-white">
                                {initials}
                            </div>
                            <span className="text-sm text-gray-300 hidden sm:block max-w-[150px] truncate">{user.email}</span>
                            <ChevronDown className="h-3.5 w-3.5 text-gray-400" />
                        </button>

                        {/* Profile Dropdown */}
                        {showProfileMenu && (
                            <div className="absolute right-0 top-full mt-2 w-64 bg-[#0a0a0a] border border-white/10 rounded-xl shadow-2xl overflow-hidden animate-in fade-in slide-in-from-top-1 z-50">
                                <div className="px-4 py-3 border-b border-white/5">
                                    <p className="text-sm font-medium text-white truncate">{user.email}</p>
                                    <p className="text-xs text-gray-500 mt-0.5">Signed in</p>
                                </div>
                                <div className="p-1.5">
                                    <button
                                        onClick={() => { setShowProfileMenu(false); setShowSettingsModal(true) }}
                                        className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-gray-300 hover:bg-white/5 hover:text-white transition-colors"
                                    >
                                        <Settings className="h-4 w-4" />
                                        API Configuration
                                    </button>
                                    <button
                                        onClick={() => { setShowProfileMenu(false); setShowFeedbackModal(true); setFeedbackSent(false); setFeedbackText(""); setFeedbackImage(null); }}
                                        className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-gray-300 hover:bg-white/5 hover:text-white transition-colors"
                                    >
                                        <MessageSquare className="h-4 w-4" />
                                        Send Feedback
                                    </button>
                                    <button
                                        onClick={() => { setShowProfileMenu(false); setShowHelpModal(true) }}
                                        className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-gray-300 hover:bg-white/5 hover:text-white transition-colors"
                                    >
                                        <HelpCircle className="h-4 w-4" />
                                        How to Use
                                    </button>
                                    <div className="border-t border-white/5 my-1"></div>
                                    <button
                                        onClick={handleLogout}
                                        className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-red-400 hover:bg-red-500/10 hover:text-red-300 transition-colors"
                                    >
                                        <LogOut className="h-4 w-4" />
                                        Log out
                                    </button>
                                </div>
                            </div>
                        )}
                    </div>
                </div>
            </header>

            {/* ─── Settings Modal / Popup ─── */}
            {showSettingsModal && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 backdrop-blur-sm" onClick={() => setShowSettingsModal(false)}>
                    <div
                        className="bg-[#0a0a0a] border border-white/10 rounded-2xl w-full max-w-lg mx-4 shadow-2xl"
                        onClick={(e) => e.stopPropagation()}
                    >
                        {/* Modal Header */}
                        <div className="flex items-center justify-between px-6 py-4 border-b border-white/5">
                            <div className="flex items-center gap-3">
                                <div className="h-8 w-8 rounded-lg bg-white/5 flex items-center justify-center">
                                    <Settings className="h-4 w-4 text-gray-400" />
                                </div>
                                <div>
                                    <h2 className="text-base font-semibold text-white">API Configuration</h2>
                                    <p className="text-xs text-gray-500">Connect to Ollama Cloud</p>
                                </div>
                            </div>
                            <button onClick={() => setShowSettingsModal(false)} className="p-1.5 rounded-lg hover:bg-white/5 transition-colors">
                                <X className="h-4 w-4 text-gray-500" />
                            </button>
                        </div>

                        {/* Modal Body */}
                        <div className="px-6 py-5 space-y-5">
                            <div className="space-y-2">
                                <Label className="text-gray-300 text-sm">Host URL</Label>
                                <Input
                                    value={hostUrl}
                                    onChange={e => setHostUrl(e.target.value)}
                                    className="bg-white/[0.03] border-white/10 font-mono text-sm h-11"
                                />
                            </div>
                            <div className="space-y-2">
                                <Label className="text-gray-300 text-sm">API Key(s)</Label>
                                <textarea
                                    value={apiKey}
                                    onChange={e => setApiKey(e.target.value)}
                                    placeholder={"Paste one or more API keys, separated by commas.\nExample: key1, key2, key3"}
                                    className="flex w-full rounded-md border border-white/10 bg-white/[0.03] px-3 py-2 font-mono text-sm h-20 focus:outline-none focus:ring-2 focus:ring-white/20 resize-none text-white placeholder:text-gray-600"
                                />
                                <p className="text-xs text-gray-500">Multiple keys enable automatic failover if one key's quota is exceeded.</p>
                            </div>

                            {availableModels.length > 0 && (
                                <>
                                <div className="space-y-2">
                                    <Label className="text-gray-300 text-sm">Text Model</Label>
                                    <select
                                        value={selectedModel} onChange={e => setSelectedModel(e.target.value)}
                                        className="flex h-11 w-full rounded-md border border-white/10 bg-white/[0.03] px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-white/20"
                                    >
                                        {availableModels.map(m => (
                                            <option key={m} value={m} className="bg-[#0a0a0a] text-gray-200 py-1">{m}</option>
                                        ))}
                                    </select>
                                    <p className="text-xs text-gray-500">Used for text, tables, and consistency checks.</p>
                                </div>
                                <div className="space-y-2">
                                    <Label className="text-gray-300 text-sm">Vision Model (for Images/Diagrams)</Label>
                                    <select
                                        value={visionModel} onChange={e => setVisionModel(e.target.value)}
                                        className="flex h-11 w-full rounded-md border border-white/10 bg-white/[0.03] px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-white/20"
                                    >
                                        {availableModels.map(m => (
                                            <option key={m} value={m} className="bg-[#0a0a0a] text-gray-200 py-1">{m}</option>
                                        ))}
                                    </select>
                                    <p className="text-xs text-gray-500">Used for reviewing images, diagrams, and flowcharts.</p>
                                </div>
                                </>
                            )}

                            {testResult && (
                                <div className={`flex items-center gap-2 text-sm px-3 py-2.5 rounded-lg ${testResult.includes("Error") ? "bg-red-500/10 text-red-400 border border-red-500/20" : "bg-green-500/10 text-green-400 border border-green-500/20"}`}>
                                    {!testResult.includes("Error") && <CheckCircle2 className="h-4 w-4 shrink-0" />}
                                    {testResult.includes("Error") && <AlertTriangle className="h-4 w-4 shrink-0" />}
                                    {testResult}
                                </div>
                            )}

                            {/* Notice */}
                            <div className="bg-white/[0.02] border border-white/5 rounded-xl p-4 space-y-2">
                                <div className="flex gap-2.5 text-xs text-gray-500">
                                    <AlertTriangle className="h-3.5 w-3.5 text-amber-500/70 shrink-0 mt-0.5" />
                                    <p>In case quota is exceeded, use another account to login to Ollama, generate a new key, and try again.</p>
                                </div>
                                <a href="https://ollama.com/settings/keys" target="_blank" className="inline-flex items-center gap-1.5 text-xs text-gray-400 hover:text-white transition-colors">
                                    <ExternalLink className="h-3 w-3" />
                                    Generate API Key on ollama.com
                                </a>
                            </div>

                            {/* How To Use Section */}
                            <div className="pt-4 border-t border-white/5 space-y-3">
                                <h3 className="text-sm font-semibold text-white">How to Use DOC-AI</h3>
                                <div className="space-y-4 text-xs text-gray-400">
                                    <div className="flex gap-3">
                                        <div className="h-5 w-5 rounded-full bg-white/10 text-white flex items-center justify-center shrink-0">1</div>
                                        <p><span className="text-white font-medium">Get your API Key:</span> Log into your Ollama account using the link above. Go to Settings &gt; API Keys, generate a new key, and paste it in the field above.</p>
                                    </div>
                                    <div className="flex gap-3">
                                        <div className="h-5 w-5 rounded-full bg-white/10 text-white flex items-center justify-center shrink-0">2</div>
                                        <p><span className="text-white font-medium">Test &amp; Save:</span> Click "Test &amp; Save" below. If successful, your key is locked in and you can close this window.</p>
                                    </div>
                                    <div className="flex gap-3">
                                        <div className="h-5 w-5 rounded-full bg-white/10 text-white flex items-center justify-center shrink-0">3</div>
                                        <p><span className="text-white font-medium">Upload Document:</span> On the dashboard, upload any Word Document (.docx) and click "Start AI Review."</p>
                                    </div>
                                    <div className="flex gap-3">
                                        <div className="h-5 w-5 rounded-full bg-white/10 text-white flex items-center justify-center shrink-0">4</div>
                                        <p><span className="text-white font-medium">Wait &amp; Download:</span> Processing takes 1 to 5 minutes depending on document length. Once 100% complete, click the "Download Excel Report" button to view your findings.</p>
                                    </div>
                                </div>
                            </div>
                        </div>

                        {/* Modal Footer */}
                        <div className="px-6 py-4 border-t border-white/5 flex items-center justify-end gap-3">
                            <Button variant="ghost" onClick={() => setShowSettingsModal(false)} className="text-gray-400 hover:text-white hover:bg-white/5">
                                Cancel
                            </Button>
                            <Button onClick={handleSaveSettings} disabled={savingSettings || !apiKey} className="bg-white text-black hover:bg-gray-200 px-6">
                                {savingSettings ? "Testing..." : "Test & Save"}
                            </Button>
                        </div>
                    </div>
                </div>
            )}

            {/* ─── Feedback Modal ─── */}
            {showFeedbackModal && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 backdrop-blur-sm" onClick={() => setShowFeedbackModal(false)}>
                    <div
                        className="bg-[#0a0a0a] border border-white/10 rounded-2xl w-full max-w-lg mx-4 shadow-2xl"
                        onClick={(e) => e.stopPropagation()}
                    >
                        <div className="flex items-center justify-between px-6 py-4 border-b border-white/5">
                            <div className="flex items-center gap-3">
                                <div className="h-8 w-8 rounded-lg bg-purple-500/10 flex items-center justify-center">
                                    <MessageSquare className="h-4 w-4 text-purple-400" />
                                </div>
                                <div>
                                    <h3 className="text-sm font-semibold text-white">Send Feedback</h3>
                                    <p className="text-xs text-gray-500">Let us know how we can improve</p>
                                </div>
                            </div>
                            <button onClick={() => setShowFeedbackModal(false)} className="p-1 rounded-lg hover:bg-white/5"><X className="h-4 w-4 text-gray-500" /></button>
                        </div>

                        <div className="px-6 py-5 space-y-4">
                            {feedbackSent ? (
                                <div className="text-center py-8">
                                    <CheckCircle2 className="h-12 w-12 text-green-400 mx-auto mb-3" />
                                    <h4 className="text-lg font-semibold text-white mb-1">Thank you!</h4>
                                    <p className="text-sm text-gray-400">Your feedback has been sent to the DOC-AI team.</p>
                                </div>
                            ) : (
                                <>
                                    <div>
                                        <Label className="text-gray-400 text-xs mb-2 block">Feedback Type</Label>
                                        <div className="flex gap-2">
                                            {[{ v: "bug", l: " Bug Report" }, { v: "feature", l: " Feature Request" }, { v: "general", l: " General" }].map(t => (
                                                <button key={t.v} onClick={() => setFeedbackType(t.v)}
                                                    className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-all ${feedbackType === t.v ? 'bg-purple-500/20 text-purple-400 border border-purple-500/30' : 'bg-white/5 text-gray-500 border border-white/5 hover:text-gray-300'}`}
                                                >{t.l}</button>
                                            ))}
                                        </div>
                                    </div>
                                    <div>
                                        <Label className="text-gray-400 text-xs mb-2 block">Your Feedback</Label>
                                        <textarea
                                            value={feedbackText}
                                            onChange={(e) => setFeedbackText(e.target.value)}
                                            placeholder="Describe the issue or suggestion in detail..."
                                            className="w-full h-32 bg-white/5 border border-white/10 rounded-xl px-4 py-3 text-sm text-white placeholder:text-gray-600 focus:outline-none focus:border-purple-500/50 resize-none"
                                        />
                                    </div>
                                    <div>
                                        <Label className="text-gray-400 text-xs mb-2 block">Attach Screenshot (optional)</Label>
                                        {feedbackImage ? (
                                            <div className="relative border border-white/10 rounded-xl overflow-hidden bg-white/5">
                                                <img src={URL.createObjectURL(feedbackImage)} alt="Preview" className="max-h-40 w-full object-contain" />
                                                <button
                                                    onClick={() => setFeedbackImage(null)}
                                                    className="absolute top-2 right-2 bg-black/70 rounded-full p-1 hover:bg-black"
                                                >
                                                    <X className="h-3 w-3 text-white" />
                                                </button>
                                            </div>
                                        ) : (
                                            <button
                                                onClick={() => feedbackImageRef.current?.click()}
                                                className="w-full border border-dashed border-white/10 rounded-xl py-4 flex flex-col items-center gap-1 text-gray-500 hover:border-purple-500/30 hover:text-purple-400 transition-colors"
                                            >
                                                <UploadCloud className="h-5 w-5" />
                                                <span className="text-xs">Click to attach an image</span>
                                            </button>
                                        )}
                                        <input
                                            type="file"
                                            ref={feedbackImageRef}
                                            accept="image/*"
                                            className="hidden"
                                            onChange={(e) => { if (e.target.files?.[0]) setFeedbackImage(e.target.files[0]) }}
                                        />
                                    </div>
                                </>
                            )}
                        </div>

                        {!feedbackSent && (
                            <div className="px-6 py-4 border-t border-white/5 flex items-center justify-end gap-3">
                                <Button variant="ghost" onClick={() => setShowFeedbackModal(false)} className="text-gray-400 hover:text-white hover:bg-white/5">Cancel</Button>
                                <Button
                                    disabled={sendingFeedback || !feedbackText.trim()}
                                    onClick={async () => {
                                        setSendingFeedback(true)
                                        try {
                                            const formData = new FormData()
                                            formData.append('user_email', user.email || 'unknown')
                                            formData.append('type', feedbackType)
                                            formData.append('message', feedbackText)
                                            if (feedbackImage) formData.append('image', feedbackImage)
                                            await fetch(`${API_BASE_URL}/api/feedback`, {
                                                method: "POST",
                                                body: formData
                                            })
                                            setFeedbackSent(true)
                                        } catch { /* silent */ }
                                        setSendingFeedback(false)
                                    }}
                                    className="bg-purple-600 text-white hover:bg-purple-500 px-6 flex items-center gap-2"
                                >
                                    <Send className="h-3.5 w-3.5" />
                                    {sendingFeedback ? "Sending..." : "Send Feedback"}
                                </Button>
                            </div>
                        )}
                    </div>
                </div>
            )}

            {/* ─── How to Use Modal ─── */}
            {showHelpModal && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/60 backdrop-blur-sm" onClick={() => setShowHelpModal(false)}>
                    <div
                        className="bg-[#0a0a0a] border border-white/10 rounded-2xl w-full max-w-lg mx-4 shadow-2xl max-h-[80vh] overflow-y-auto"
                        onClick={(e) => e.stopPropagation()}
                    >
                        <div className="flex items-center justify-between px-6 py-4 border-b border-white/5 sticky top-0 bg-[#0a0a0a] z-10">
                            <div className="flex items-center gap-3">
                                <div className="h-8 w-8 rounded-lg bg-blue-500/10 flex items-center justify-center">
                                    <HelpCircle className="h-4 w-4 text-blue-400" />
                                </div>
                                <div>
                                    <h3 className="text-sm font-semibold text-white">How to Use DOC-AI</h3>
                                    <p className="text-xs text-gray-500">Quick start guide</p>
                                </div>
                            </div>
                            <button onClick={() => setShowHelpModal(false)} className="p-1 rounded-lg hover:bg-white/5"><X className="h-4 w-4 text-gray-500" /></button>
                        </div>

                        <div className="px-6 py-5 space-y-5 text-sm text-gray-400">
                            {[
                                { n: "1", title: "Configure API Key", desc: "Click your profile icon → \"API Configuration\". Paste your Ollama API Key and click \"Test & Save\"." },
                                { n: "2", title: "Select File Type", desc: "Choose between \"Word Document\" (.docx) or \"Excel Sheet\" (.xlsx) using the toggle above the upload area." },
                                { n: "3", title: "Upload Your File", desc: "Click the dashed upload box to browse your computer and select a file." },
                                { n: "4", title: "Choose Review Mode", desc: "Normal Mode: Checks grammar, terminology, and units (faster). Pro Mode: Deep 12-category technical review." },
                                { n: "5", title: "Start AI Review", desc: "Click \"Start AI Review\" and wait 1-5 minutes. A live progress bar tracks the analysis." },
                                { n: "6", title: "Download Report", desc: "Once complete, click \"Download Excel Report\" to get a professional report with findings, severity levels, and fix instructions." },
                            ].map(step => (
                                <div key={step.n} className="flex gap-3">
                                    <div className="h-6 w-6 rounded-full bg-blue-500/10 text-blue-400 flex items-center justify-center shrink-0 text-xs font-bold">{step.n}</div>
                                    <div><span className="text-white font-medium">{step.title}:</span> {step.desc}</div>
                                </div>
                            ))}
                            <div className="border-t border-white/5 pt-4 text-xs text-gray-500">
                                For issues or questions, use the "Send Feedback" option in your profile menu.
                            </div>
                        </div>
                    </div>
                </div>
            )}

            {/* ─── Main Content ─── */}
            <main className="flex-1 flex flex-col items-center justify-start px-6 py-12 max-w-4xl mx-auto w-full">

                {/* Upload Section */}
                <div className="w-full flex flex-col md:flex-row md:items-end justify-between gap-4 mb-8">
                    <div className="space-y-2">
                        <h2 className="text-2xl font-bold tracking-tight text-white">Document Review</h2>
                        <p className="text-sm text-gray-500">
                            Upload a {fileType === 'excel' ? 'Excel sheet' : 'Word document'} and get an AI review.
                        </p>
                    </div>

                    <div className="flex bg-white/5 p-1 rounded-xl border border-white/5">
                        <button
                            onClick={() => { setFileType('doc'); setSelectedFile(null); }}
                            className={`px-4 py-2 rounded-lg text-xs font-medium transition-all ${fileType === 'doc' ? 'bg-blue-500/20 text-blue-400 border border-blue-500/30' : 'text-gray-500 hover:text-gray-300'}`}
                        >
                            Word Document
                        </button>
                        <button
                            onClick={() => { setFileType('excel'); setSelectedFile(null); }}
                            className={`px-4 py-2 rounded-lg text-xs font-medium transition-all ${fileType === 'excel' ? 'bg-green-500/20 text-green-400 border border-green-500/30' : 'text-gray-500 hover:text-gray-300'}`}
                        >
                            Excel Sheet
                        </button>
                    </div>
                </div>

                <Card className="w-full border-white/10 bg-white/[0.02] shadow-none">
                    <CardContent className="p-8 space-y-6">

                        {!selectedFile ? (
                            <div
                                className="border-2 border-dashed border-white/10 rounded-2xl p-16 flex flex-col items-center justify-center text-center hover:border-white/20 hover:bg-white/[0.02] transition-all duration-300 cursor-pointer group"
                                onClick={() => fileInputRef.current?.click()}
                            >
                                <div className="h-16 w-16 rounded-2xl bg-white/5 flex items-center justify-center mb-5 group-hover:bg-white/10 transition-colors">
                                    <UploadCloud className="h-7 w-7 text-gray-500 group-hover:text-gray-300 transition-colors" />
                                </div>
                                <h3 className="text-lg font-medium text-gray-200 mb-1">Click to upload document</h3>
                                <p className="text-sm text-gray-500">Supports {fileType === 'excel' ? '.xlsx, .xls' : '.docx, .doc'}</p>
                                <input
                                    type="file"
                                    ref={fileInputRef}
                                    onChange={handleFileChange}
                                    accept={fileType === 'excel' ? ".xlsx,.xls" : ".docx,.doc"}
                                    className="hidden"
                                />
                            </div>
                        ) : (
                            <div className="border border-white/10 p-5 rounded-xl flex items-center justify-between bg-white/[0.02]">
                                <div className="flex items-center gap-4">
                                    <div className={`h-12 w-12 ${fileType === 'excel' ? 'bg-green-500/10 text-green-400' : 'bg-blue-500/10 text-blue-400'} rounded-xl flex items-center justify-center`}>
                                        <FileText className="h-5 w-5" />
                                    </div>
                                    <div>
                                        <h4 className="font-medium text-sm text-gray-200">{selectedFile.name}</h4>
                                        <p className="text-xs text-gray-500 mt-0.5">{(selectedFile.size / 1024 / 1024).toFixed(2)} MB</p>
                                    </div>
                                </div>
                                <Button variant="ghost" size="sm" onClick={() => { setSelectedFile(null); setFindings([]); setReportUrl(null); setProgressPct(0); setProgressMsg("") }} disabled={reviewing} className="text-gray-400 hover:text-white">
                                    Remove
                                </Button>
                            </div>
                        )}

                        {/* Progress */}
                        {(reviewing || progressPct > 0) && (
                            <div className="space-y-3">
                                <div className="flex justify-between text-sm">
                                    <span className="text-gray-300">{progressMsg || "Processing..."}</span>
                                    <span className="text-gray-500 tabular-nums">{progressPct}%</span>
                                </div>
                                <div className="h-1.5 w-full bg-white/5 rounded-full overflow-hidden">
                                    <div
                                        className="h-full bg-white rounded-full transition-all duration-500 ease-out"
                                        style={{ width: `${progressPct}%` }}
                                    ></div>
                                </div>
                            </div>
                        )}

                        {/* Actions */}
                        <div className="flex flex-wrap items-center gap-4 pt-2">
                            {/* Mode Toggle */}
                            <div className="flex bg-white/5 p-1 rounded-lg border border-white/10 mr-2">
                                <button
                                    onClick={() => setReviewMode("normal")}
                                    className={`px-4 py-1.5 text-xs font-medium rounded-md transition-all ${reviewMode === "normal" ? "bg-white text-black shadow-lg" : "text-gray-400 hover:text-white"}`}
                                >
                                    Normal
                                </button>
                                <button
                                    onClick={() => setReviewMode("pro")}
                                    className={`px-4 py-1.5 text-xs font-medium rounded-md transition-all ${reviewMode === "pro" ? "bg-white text-black shadow-lg" : "text-gray-400 hover:text-white"}`}
                                >
                                    Pro
                                </button>
                            </div>

                            <Button
                                onClick={handleStartReview}
                                disabled={reviewing || !selectedFile}
                                className="bg-white text-black hover:bg-gray-200 px-8 h-11 text-sm font-medium"
                            >
                                {reviewing ? "Reviewing..." : "Start AI Review"}
                            </Button>
                            {reportUrl && (
                                <Button asChild variant="outline" className="border-white/15 bg-transparent text-white hover:bg-white/5 h-11">
                                    <a href={reportUrl} download>Download Report (.xlsx)</a>
                                </Button>
                            )}
                        </div>
                    </CardContent>
                </Card>

                {/* Findings */}
                {findings.length > 0 && (
                    <div className="w-full mt-10 space-y-4">
                        <div className="flex items-center justify-between">
                            <h3 className="text-lg font-semibold text-white">Review Findings</h3>
                            <span className="text-xs text-gray-500 bg-white/5 px-3 py-1 rounded-full">{findings.length} issues found</span>
                        </div>
                        <div className="space-y-3 max-h-[500px] overflow-y-auto pr-1">
                            {findings.map((f, i) => (
                                <div key={i} className="bg-white/[0.02] border border-white/5 rounded-xl p-4 flex gap-4 text-sm hover:border-white/10 transition-colors">
                                    <div className="min-w-fit pt-0.5">
                                        {f.severity === 'CRITICAL' && <span className="px-2.5 py-1 bg-red-500/15 text-red-400 rounded-md text-xs font-medium">Critical</span>}
                                        {f.severity === 'MAJOR' && <span className="px-2.5 py-1 bg-orange-500/15 text-orange-400 rounded-md text-xs font-medium">Major</span>}
                                        {f.severity === 'MINOR' && <span className="px-2.5 py-1 bg-yellow-500/15 text-yellow-400 rounded-md text-xs font-medium">Minor</span>}
                                    </div>
                                    <div className="space-y-1 w-full">
                                        <div className="flex items-center justify-between text-gray-500 text-xs">
                                            <span>{f.category?.replace(/_/g, " ")} &bull; Page {f.page}</span>
                                            <span className="truncate max-w-[200px]">{f.section}</span>
                                        </div>
                                        <p className="text-gray-200 leading-relaxed">{f.comment}</p>
                                    </div>
                                </div>
                            ))}
                        </div>
                    </div>
                )}
            </main>
        </div>
    )
}
