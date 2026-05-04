"use client"

import type React from "react"

import { useState } from "react"
import Link from "next/link"
import { useRouter } from "next/navigation"
import { Calendar, Eye, EyeOff, ArrowLeft, Shield, User } from "lucide-react"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { ThemeToggle } from "@/components/theme-toggle"
import { useToast } from "@/hooks/use-toast"
import { API_BASE_URL } from "@/lib/config"

export default function LoginPage() {
  const [showPassword, setShowPassword] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const [loginType, setLoginType] = useState<"user" | "admin">("user")
  const [formData, setFormData] = useState({
    email: "",
    password: "",
  })
  const { toast } = useToast()
  const router = useRouter()

  const handleInputChange = (field: string, value: string) => {
    setFormData((prev) => ({
      ...prev,
      [field]: value,
    }))
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setIsLoading(true)

    // Add password validation before API call
    if (formData.password.length < 8) {
      toast({
        variant: "destructive",
        title: "Password Too Short",
        description: "Password must be at least 8 characters long.",
      })
      setIsLoading(false)
      return
    }

    // Determine API endpoint based on login type
    const apiEndpoint =
      loginType === "admin"
        ? `${API_BASE_URL}/admin-login/`
        : `${API_BASE_URL}/login/`

    try {
      const response = await fetch(apiEndpoint, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email: formData.email,
          password: formData.password,
        }),
      })

      const data = await response.json()

      if (response.ok) {
        // API returns user data directly in the response
        const userInfo = {
          email: data.email,
          username: data.username,
          role: data.role,
        }

        // Store user info and login status
        localStorage.setItem("user", JSON.stringify(userInfo))
        localStorage.setItem("isLoggedIn", "true")

        toast({
          variant: "success",
          title: "Login Successful!",
          description: `Welcome back, ${data.username}! Redirecting to dashboard...`,
        })

        // Small delay to show success message, then redirect
        setTimeout(() => {
          if (loginType === "admin" || data.role === "admin") {
            router.push("/admin/dashboard")
          } else {
            router.push("/dashboard")
          }
        }, 1000)
      } else {
        toast({
          variant: "destructive",
          title: "Login Failed",
          description: data.detail || data.message || "Invalid credentials. Please check your email and password.",
        })
      }
    } catch (error) {
      console.error("Login error:", error)
      toast({
        variant: "destructive",
        title: "Network Error",
        description: "Please check your internet connection and try again.",
      })
    } finally {
      setIsLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50 dark:from-gray-900 dark:via-gray-800 dark:to-blue-900 flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        {/* Header */}
        <div className="flex items-center justify-between mb-6 sm:mb-8">
          <Link
            href="/"
            className="inline-flex items-center text-blue-600 dark:text-blue-400 hover:text-blue-700 dark:hover:text-blue-300 transition-colors duration-200"
          >
            <ArrowLeft className="h-4 w-4 mr-2" />
            <span className="hidden sm:inline">Back to Home</span>
            <span className="sm:hidden">Back</span>
          </Link>
          <ThemeToggle />
        </div>

        <Card className="shadow-xl border-0 bg-white/90 dark:bg-gray-800/90 backdrop-blur-sm">
          <CardHeader className="text-center pb-6 sm:pb-8 px-4 sm:px-6">
            <div className="flex items-center justify-center space-x-2 mb-3 sm:mb-4">
              <Calendar className="h-8 w-8 sm:h-10 sm:w-10 text-blue-600 dark:text-blue-400" />
              <span className="text-xl sm:text-2xl font-bold text-gray-900 dark:text-white">
                <span className="hidden sm:inline">Smart Scheduler</span>
                <span className="sm:hidden">Smart</span>
              </span>
            </div>
            <CardTitle className="text-2xl sm:text-3xl font-bold text-gray-900 dark:text-white">Welcome Back</CardTitle>
            <CardDescription className="text-gray-600 dark:text-gray-300 text-base sm:text-lg">
              Sign in to your account
            </CardDescription>
          </CardHeader>

          <CardContent className="px-4 sm:px-6">
            <Tabs value={loginType} onValueChange={(value) => setLoginType(value as "user" | "admin")} className="mb-6">
              <TabsList className="grid w-full grid-cols-2">
                <TabsTrigger value="user" className="flex items-center space-x-2">
                  <User className="h-4 w-4" />
                  <span>User</span>
                </TabsTrigger>
                <TabsTrigger value="admin" className="flex items-center space-x-2">
                  <Shield className="h-4 w-4" />
                  <span>Admin</span>
                </TabsTrigger>
              </TabsList>

              <TabsContent value="user" className="mt-6">
                <form onSubmit={handleSubmit} className="space-y-4 sm:space-y-6">
                  <div className="space-y-2">
                    <Label htmlFor="email" className="text-sm font-medium text-gray-700 dark:text-gray-300">
                      Email Address
                    </Label>
                    <Input
                      id="email"
                      type="email"
                      placeholder="Enter your email"
                      value={formData.email}
                      onChange={(e) => handleInputChange("email", e.target.value)}
                      required
                      className="h-11 sm:h-12 text-base transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50 backdrop-blur-sm"
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="password" className="text-sm font-medium text-gray-700 dark:text-gray-300">
                      Password
                    </Label>
                    <div className="relative">
                      <Input
                        id="password"
                        type={showPassword ? "text" : "password"}
                        placeholder="Enter your password (min 8 characters)"
                        value={formData.password}
                        onChange={(e) => handleInputChange("password", e.target.value)}
                        required
                        minLength={8}
                        className="h-11 sm:h-12 text-base pr-10 transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50 backdrop-blur-sm"
                      />
                      <button
                        type="button"
                        onClick={() => setShowPassword(!showPassword)}
                        className="absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 transition-colors duration-200"
                      >
                        {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                      </button>
                    </div>
                  </div>

                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 sm:gap-0">
                    <label className="flex items-center">
                      <input type="checkbox" className="rounded border-gray-300 text-blue-600 focus:ring-blue-500" />
                      <span className="ml-2 text-sm text-gray-600 dark:text-gray-300">Remember me</span>
                    </label>
                    <Link
                      href="#"
                      className="text-sm text-blue-600 dark:text-blue-400 hover:text-blue-700 dark:hover:text-blue-300 transition-colors duration-200"
                    >
                      Forgot password?
                    </Link>
                  </div>

                  <Button
                    type="submit"
                    className="w-full h-11 sm:h-12 text-base bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 dark:from-blue-500 dark:to-indigo-500 dark:hover:from-blue-600 dark:hover:to-indigo-600 shadow-lg hover:shadow-xl transition-all duration-200"
                    disabled={isLoading}
                  >
                    {isLoading ? (
                      <div className="flex items-center">
                        <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                        Signing In...
                      </div>
                    ) : (
                      "Sign In"
                    )}
                  </Button>
                </form>
              </TabsContent>

              <TabsContent value="admin" className="mt-6">
                <form onSubmit={handleSubmit} className="space-y-4 sm:space-y-6">
                  <div className="space-y-2">
                    <Label htmlFor="admin-email" className="text-sm font-medium text-gray-700 dark:text-gray-300">
                      Admin Email
                    </Label>
                    <Input
                      id="admin-email"
                      type="email"
                      placeholder="Enter admin email"
                      value={formData.email}
                      onChange={(e) => handleInputChange("email", e.target.value)}
                      required
                      className="h-11 sm:h-12 text-base transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50 backdrop-blur-sm"
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="admin-password" className="text-sm font-medium text-gray-700 dark:text-gray-300">
                      Admin Password
                    </Label>
                    <div className="relative">
                      <Input
                        id="admin-password"
                        type={showPassword ? "text" : "password"}
                        placeholder="Enter admin password (min 8 characters)"
                        value={formData.password}
                        onChange={(e) => handleInputChange("password", e.target.value)}
                        required
                        minLength={8}
                        className="h-11 sm:h-12 text-base pr-10 transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50 backdrop-blur-sm"
                      />
                      <button
                        type="button"
                        onClick={() => setShowPassword(!showPassword)}
                        className="absolute right-3 top-1/2 transform -translate-y-1/2 text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 transition-colors duration-200"
                      >
                        {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
                      </button>
                    </div>
                  </div>

                  <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-3">
                    <p className="text-sm text-amber-800 dark:text-amber-200">
                      <Shield className="h-4 w-4 inline mr-1" />
                      Admin access required for administrative functions
                    </p>
                  </div>

                  <Button
                    type="submit"
                    className="w-full h-11 sm:h-12 text-base bg-gradient-to-r from-amber-600 to-orange-600 hover:from-amber-700 hover:to-orange-700 dark:from-amber-500 dark:to-orange-500 dark:hover:from-amber-600 dark:hover:to-orange-600 shadow-lg hover:shadow-xl transition-all duration-200"
                    disabled={isLoading}
                  >
                    {isLoading ? (
                      <div className="flex items-center">
                        <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                        Signing In...
                      </div>
                    ) : (
                      "Admin Sign In"
                    )}
                  </Button>
                </form>
              </TabsContent>
            </Tabs>

            <div className="mt-4 sm:mt-6 text-center">
              <p className="text-sm text-gray-600 dark:text-gray-300">
                {"Don't have an account? "}
                <Link
                  href="/signup"
                  className="text-blue-600 dark:text-blue-400 hover:text-blue-700 dark:hover:text-blue-300 font-medium transition-colors duration-200"
                >
                  Sign up here
                </Link>
              </p>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  )
}
