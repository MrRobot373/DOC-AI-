import { Component, type ErrorInfo, type ReactNode } from "react"

interface Props {
    children: ReactNode
    /** Optional fallback UI; defaults to a small inline notice. */
    fallback?: ReactNode
}
interface State {
    hasError: boolean
    message?: string
}

/**
 * Catches render-time errors in its subtree so one broken component (e.g. the
 * document viewer) can never blank the whole dashboard.
 */
export default class ErrorBoundary extends Component<Props, State> {
    state: State = { hasError: false }

    static getDerivedStateFromError(err: Error): State {
        return { hasError: true, message: err?.message }
    }

    componentDidCatch(error: Error, info: ErrorInfo) {
        console.error("ErrorBoundary caught:", error, info)
    }

    render() {
        if (this.state.hasError) {
            return (
                this.props.fallback ?? (
                    <div className="rounded-xl border border-amber-500/30 bg-amber-500/10 px-4 py-3 text-sm text-amber-300">
                        This panel hit an error and was hidden so the rest of the page keeps working.
                        {this.state.message ? <span className="block text-xs text-amber-400/70 mt-1">{this.state.message}</span> : null}
                    </div>
                )
            )
        }
        return this.props.children
    }
}
