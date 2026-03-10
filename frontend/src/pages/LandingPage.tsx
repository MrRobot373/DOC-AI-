import { Link } from "react-router-dom"
import { Button } from "@/components/ui/button"

export default function LandingPage() {
    return (
        <div className="flex flex-col min-h-screen bg-black text-white selection:bg-white/30">
            <header className="absolute inset-x-0 top-0 z-50 flex h-24 items-center justify-between px-8">
                <div className="flex items-center gap-6">
                    <h1 className="text-2xl font-bold tracking-tight">DOC-AI</h1>
                    <div className="h-6 w-[1px] bg-white/20 hidden sm:block"></div>
                    {/* Using text for GMS logo if img isn't available, but keeping slot ready */}
                    <div className="flex items-center gap-2">
                        <span className="font-semibold text-red-500 tracking-wider">GMS</span>
                        <span className="text-xs text-gray-400 hidden sm:block uppercase tracking-widest mt-0.5">The Technology Company</span>
                    </div>
                </div>
                <div className="flex items-center gap-4">
                    <Button variant="ghost" asChild className="text-gray-300 hover:text-white hover:bg-white/10">
                        <Link to="/login">Log in</Link>
                    </Button>
                    <Button asChild className="bg-white text-black hover:bg-gray-200 rounded-full px-6">
                        <Link to="/signup">Sign up</Link>
                    </Button>
                </div>
            </header>

            <main className="flex-1 flex flex-col items-center justify-center relative overflow-hidden px-6">
                {/* Subtle background glow */}
                <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[800px] h-[400px] bg-white/5 rounded-[100%] blur-[120px] pointer-events-none"></div>

                <div className="z-10 text-center max-w-3xl space-y-8">
                    <h2 className="text-5xl sm:text-7xl font-bold tracking-tighter leading-tight bg-gradient-to-b from-white to-gray-500 bg-clip-text text-transparent">
                        Intelligent Document Review for GMS Engineers.
                    </h2>
                    <p className="text-lg sm:text-xl text-gray-400 max-w-2xl mx-auto font-light leading-relaxed">
                        Automate your design specifications and test report reviews using powerful AI models. Connect your Ollama instance and streamline your workflow in seconds.
                    </p>
                    <div className="pt-8 flex items-center justify-center">
                        <Button size="lg" asChild className="bg-white text-black hover:bg-gray-200 rounded-full px-8 text-base">
                            <Link to="/signup">Get Started</Link>
                        </Button>
                    </div>
                </div>
            </main>
        </div>
    )
}
