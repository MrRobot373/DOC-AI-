import { useState } from "react"
import { Link, useNavigate } from "react-router-dom"
import { supabase } from "@/lib/supabase"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card"

export default function Signup() {
    const [email, setEmail] = useState("")
    const [password, setPassword] = useState("")
    const [loading, setLoading] = useState(false)
    const [error, setError] = useState<string | null>(null)
    const [success, setSuccess] = useState(false)

    const navigate = useNavigate()

    const handleSignup = async (e: React.FormEvent) => {
        e.preventDefault()
        setLoading(true)
        setError(null)

        const { error: signUpError } = await supabase.auth.signUp({
            email,
            password,
        })

        if (signUpError) {
            setError(signUpError.message)
            setLoading(false)
        } else {
            setSuccess(true)
            setLoading(false)
            // Small delay before redirecting to dashboard or login
            setTimeout(() => {
                navigate('/dashboard')
            }, 2000)
        }
    }

    return (
        <div className="flex min-h-screen items-center justify-center bg-black p-4">
            <Link to="/" className="absolute top-8 left-8 text-white font-bold text-xl tracking-tight">DOC-AI</Link>

            <Card className="w-full max-w-md border-white/10 bg-black text-white shadow-2xl">
                <CardHeader className="space-y-2 text-center">
                    <CardTitle className="text-2xl tracking-tight">Create an account</CardTitle>
                    <CardDescription className="text-gray-400">
                        Sign up to start reviewing your documents
                    </CardDescription>
                </CardHeader>
                <form onSubmit={handleSignup}>
                    <CardContent className="space-y-4">
                        {error && (
                            <div className="p-3 text-sm bg-red-950/50 border border-red-500/50 text-red-200 rounded-md">
                                {error}
                            </div>
                        )}
                        {success && (
                            <div className="p-3 text-sm bg-green-950/50 border border-green-500/50 text-green-200 rounded-md">
                                Account created successfully! redirecting...
                            </div>
                        )}
                        <div className="space-y-2">
                            <Label htmlFor="email" className="text-gray-300">Email</Label>
                            <Input
                                id="email"
                                type="email"
                                placeholder="name@company.com"
                                value={email}
                                onChange={(e) => setEmail(e.target.value)}
                                required
                                className="bg-white/5 border-white/10 text-white placeholder:text-gray-500 focus-visible:ring-white/20"
                            />
                        </div>
                        <div className="space-y-2">
                            <Label htmlFor="password" className="text-gray-300">Password</Label>
                            <Input
                                id="password"
                                type="password"
                                value={password}
                                onChange={(e) => setPassword(e.target.value)}
                                required
                                minLength={6}
                                className="bg-white/5 border-white/10 text-white placeholder:text-gray-500 focus-visible:ring-white/20"
                            />
                        </div>
                    </CardContent>
                    <CardFooter className="flex flex-col space-y-4">
                        <Button
                            type="submit"
                            className="w-full bg-white text-black hover:bg-gray-200"
                            disabled={loading || success}
                        >
                            {loading ? "Creating account..." : "Sign Up"}
                        </Button>
                        <div className="text-center text-sm text-gray-400">
                            Already have an account?{" "}
                            <Link to="/login" className="text-white hover:underline underline-offset-4">
                                Sign in
                            </Link>
                        </div>
                    </CardFooter>
                </form>
            </Card>
        </div>
    )
}
