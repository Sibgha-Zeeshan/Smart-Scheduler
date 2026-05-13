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
  <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50 dark:from-gray-900 dark:via-gray-800 dark:to-blue-900 flex items-center justify-center p-4 py-8">
    <div className="w-full max-w-lg">
      {/* Header */}
      <div className="flex items-center justify-between mb-5">
        <Link
          href="/"
          className="inline-flex items-center text-blue-600 dark:text-blue-400 hover:text-blue-700 dark:hover:text-blue-300 transition-colors duration-200 text-sm font-medium"
        >
          <ArrowLeft className="h-4 w-4 mr-1.5" />
          Back to Home
        </Link>
        <ThemeToggle />
      </div>

      <Card className="shadow-xl border-0 bg-white/90 dark:bg-gray-800/90 backdrop-blur-sm">
        <CardHeader className="text-center pt-7 pb-5 px-6">
          <div className="flex items-center justify-center space-x-2 mb-3">
            <Calendar className="h-8 w-8 text-blue-600 dark:text-blue-400" />
            <span className="text-xl font-bold text-gray-900 dark:text-white">
              Smart Scheduler
            </span>
          </div>

          <CardTitle className="text-2xl font-bold text-gray-900 dark:text-white">
            Welcome Back
          </CardTitle>

          <CardDescription className="text-gray-600 dark:text-gray-300 text-sm mt-1">
            Sign in to your account
          </CardDescription>
        </CardHeader>

        <CardContent className="px-6 pb-7">
          <Tabs
            value={loginType}
            onValueChange={(value) => setLoginType(value as "user" | "admin")}
            className="mb-5"
          >
            <TabsList className="grid w-full grid-cols-2 h-11">
              <TabsTrigger
                value="user"
                className="flex items-center space-x-2 text-sm"
              >
                <User className="h-4 w-4" />
                <span>User</span>
              </TabsTrigger>

              <TabsTrigger
                value="admin"
                className="flex items-center space-x-2 text-sm"
              >
                <Shield className="h-4 w-4" />
                <span>Admin</span>
              </TabsTrigger>
            </TabsList>

            {/* USER TAB */}
            <TabsContent value="user" className="mt-5">
              <form onSubmit={handleSubmit} className="space-y-4">

                {/* Email */}
                <div className="space-y-1.5">
                  <Label
                    htmlFor="email"
                    className="text-sm font-medium text-gray-700 dark:text-gray-300"
                  >
                    Email Address
                  </Label>

                  <Input
                    id="email"
                    type="email"
                    placeholder="Enter your email"
                    value={formData.email}
                    onChange={(e) =>
                      handleInputChange("email", e.target.value)
                    }
                    required
                    className="h-11 text-sm transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50"
                  />
                </div>

                {/* Password */}
                <div className="space-y-1.5">
                  <Label
                    htmlFor="password"
                    className="text-sm font-medium text-gray-700 dark:text-gray-300"
                  >
                    Password
                  </Label>

                  <div className="relative">
                    <Input
                      id="password"
                      type={showPassword ? "text" : "password"}
                      placeholder="Enter your password (min 8 characters)"
                      value={formData.password}
                      onChange={(e) =>
                        handleInputChange("password", e.target.value)
                      }
                      required
                      minLength={8}
                      className="h-11 text-sm pr-10 transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50"
                    />

                    <button
                      type="button"
                      onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 transition-colors duration-200"
                    >
                      {showPassword ? (
                        <EyeOff className="h-4 w-4" />
                      ) : (
                        <Eye className="h-4 w-4" />
                      )}
                    </button>
                  </div>

                  <p className="text-xs text-gray-500 dark:text-gray-400">
                    Must be at least 8 characters
                  </p>
                </div>

                {/* Remember + Forgot */}
                <div className="flex items-center justify-between">
                  <label className="flex items-center">
                    <input
                      type="checkbox"
                      className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                    />
                    <span className="ml-2 text-sm text-gray-600 dark:text-gray-300">
                      Remember me
                    </span>
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
                  className="w-full h-11 text-sm font-semibold bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 dark:from-blue-500 dark:to-indigo-500 dark:hover:from-blue-600 dark:hover:to-indigo-600 shadow-lg hover:shadow-xl transition-all duration-200 mt-2"
                  disabled={isLoading}
                >
                  {isLoading ? (
                    <div className="flex items-center justify-center">
                      <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                      Signing In...
                    </div>
                  ) : (
                    "Sign In"
                  )}
                </Button>
              </form>
            </TabsContent>

            {/* ADMIN TAB */}
            <TabsContent value="admin" className="mt-5">
              <form onSubmit={handleSubmit} className="space-y-4">

                {/* Admin Email */}
                <div className="space-y-1.5">
                  <Label
                    htmlFor="admin-email"
                    className="text-sm font-medium text-gray-700 dark:text-gray-300"
                  >
                    Admin Email
                  </Label>

                  <Input
                    id="admin-email"
                    type="email"
                    placeholder="Enter admin email"
                    value={formData.email}
                    onChange={(e) =>
                      handleInputChange("email", e.target.value)
                    }
                    required
                    className="h-11 text-sm transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50"
                  />
                </div>

                {/* Admin Password */}
                <div className="space-y-1.5">
                  <Label
                    htmlFor="admin-password"
                    className="text-sm font-medium text-gray-700 dark:text-gray-300"
                  >
                    Admin Password
                  </Label>

                  <div className="relative">
                    <Input
                      id="admin-password"
                      type={showPassword ? "text" : "password"}
                      placeholder="Enter admin password (min 8 characters)"
                      value={formData.password}
                      onChange={(e) =>
                        handleInputChange("password", e.target.value)
                      }
                      required
                      minLength={8}
                      className="h-11 text-sm pr-10 transition-all duration-200 focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white/50 dark:bg-gray-700/50"
                    />

                    <button
                      type="button"
                      onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-700 dark:text-gray-400 dark:hover:text-gray-200 transition-colors duration-200"
                    >
                      {showPassword ? (
                        <EyeOff className="h-4 w-4" />
                      ) : (
                        <Eye className="h-4 w-4" />
                      )}
                    </button>
                  </div>

                  <p className="text-xs text-gray-500 dark:text-gray-400">
                    Must be at least 8 characters
                  </p>
                </div>

                <div className="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg p-3">
                  <p className="text-sm text-amber-800 dark:text-amber-200">
                    <Shield className="h-4 w-4 inline mr-1" />
                    Admin access required for administrative functions
                  </p>
                </div>

                <Button
                  type="submit"
                  className="w-full h-11 text-sm font-semibold bg-gradient-to-r from-amber-600 to-orange-600 hover:from-amber-700 hover:to-orange-700 dark:from-amber-500 dark:to-orange-500 dark:hover:from-amber-600 dark:hover:to-orange-600 shadow-lg hover:shadow-xl transition-all duration-200 mt-2"
                  disabled={isLoading}
                >
                  {isLoading ? (
                    <div className="flex items-center justify-center">
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

          <div className="mt-5 text-center">
            <p className="text-sm text-gray-600 dark:text-gray-300">
              Don't have an account?{" "}
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
