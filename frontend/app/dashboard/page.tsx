"use client"

import { useEffect, useState } from "react"
import { useRouter } from "next/navigation"
import { Calendar, Download, LogOut, User, BookOpen, GraduationCap, Clock, MapPin, Star } from "lucide-react"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table"
import { Badge } from "@/components/ui/badge"
import { ThemeToggle } from "@/components/theme-toggle"
import { useToast } from "@/hooks/use-toast"
import { API_BASE_URL } from "@/lib/config"

interface UserData {
  email: string
  username: string
  role: string
}

interface TimetableData {
  columns: string[]
  rows: (string | number)[][]
}

export default function DashboardPage() {
  const [userData, setUserData] = useState<UserData | null>(null)
  const [timetableData, setTimetableData] = useState<TimetableData | null>(null)
  const [latestFileName, setLatestFileName] = useState<string>("")
  const [isLoading, setIsLoading] = useState(true)
  const [isDownloading, setIsDownloading] = useState(false)
  const [mounted, setMounted] = useState(false)
  const router = useRouter()
  const { toast } = useToast()

  // Handle hydration
  useEffect(() => {
    setMounted(true)
  }, [])

  useEffect(() => {
    if (!mounted) return

    // Check if user is logged in
    try {
      const isLoggedIn = localStorage.getItem("isLoggedIn")
      const user = localStorage.getItem("user")

      console.log("Dashboard - Is logged in:", isLoggedIn)
      console.log("Dashboard - User data:", user)

      if (!isLoggedIn || isLoggedIn !== "true" || !user) {
        console.log("Dashboard - Redirecting to login: not logged in or missing user data")
        router.push("/login")
        return
      }

      const parsedUser = JSON.parse(user)
      console.log("Dashboard - Parsed user:", parsedUser)

      setUserData(parsedUser)
      fetchTimetableData()
    } catch (error) {
      console.error("Error parsing user data:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Failed to load user data. Please login again.",
      })
      router.push("/login")
    }
  }, [mounted, router, toast])

  const fetchTimetableData = async () => {
    try {
      // First, get the latest filename
      const filenameResponse = await fetch(`${API_BASE_URL}/latest-timetable-filename/`)

      if (!filenameResponse.ok) {
        throw new Error("Failed to fetch latest filename")
      }

      const filenameData = await filenameResponse.json()
      const filename = filenameData.latest_file
      setLatestFileName(filename)

      // Then, get the timetable data
      const timetableResponse = await fetch(
        `${API_BASE_URL}/timetable-data/?file_name=${filename}`,
      )

      if (!timetableResponse.ok) {
        throw new Error("Failed to fetch timetable data")
      }

      const timetableData = await timetableResponse.json()
      setTimetableData(timetableData)
    } catch (error) {
      console.error("Error fetching timetable data:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Failed to load timetable data. Please try again.",
      })
    } finally {
      setIsLoading(false)
    }
  }

  const handleDownloadTimetable = async () => {
    if (!latestFileName) return

    setIsDownloading(true)
    try {
      const response = await fetch(`${API_BASE_URL}/download/${latestFileName}`)

      if (!response.ok) {
        throw new Error("Failed to download timetable")
      }

      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement("a")
      a.style.display = "none"
      a.href = url
      a.download = latestFileName
      document.body.appendChild(a)
      a.click()
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)

      toast({
        variant: "success",
        title: "Download Complete",
        description: "Timetable has been downloaded successfully.",
      })
    } catch (error) {
      console.error("Error downloading timetable:", error)
      toast({
        variant: "destructive",
        title: "Download Failed",
        description: "Failed to download timetable. Please try again.",
      })
    } finally {
      setIsDownloading(false)
    }
  }

  const handleLogout = () => {
    if (typeof window !== "undefined") {
      localStorage.removeItem("isLoggedIn")
      localStorage.removeItem("user")
      toast({
        variant: "success",
        title: "Logged Out",
        description: "You have been logged out successfully.",
      })
      router.push("/")
    }
  }

  const getRoleIcon = (role: string) => {
    switch (role?.toLowerCase()) {
      case "student":
        return <GraduationCap className="h-5 w-5" />
      case "teacher":
        return <BookOpen className="h-5 w-5" />
      default:
        return <User className="h-5 w-5" />
    }
  }

  const getRoleColor = (role: string) => {
    switch (role?.toLowerCase()) {
      case "student":
        return "bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200"
      case "teacher":
        return "bg-green-100 text-green-800 dark:bg-green-900 dark:text-green-200"
      default:
        return "bg-gray-100 text-gray-800 dark:bg-gray-900 dark:text-gray-200"
    }
  }

  const getCourseTypeColor = (type: string) => {
    switch (type?.toLowerCase()) {
      case "core":
        return "bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200"
      case "lab":
        return "bg-purple-100 text-purple-800 dark:bg-purple-900 dark:text-purple-200"
      default:
        return "bg-gray-100 text-gray-800 dark:bg-gray-900 dark:text-gray-200"
    }
  }

  // Don't render anything until mounted (prevents hydration issues)
  if (!mounted) {
    return null
  }

  if (isLoading) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50 dark:bg-gradient-to-br dark:from-blue-900 dark:via-gray-900 dark:to-black flex items-center justify-center">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600 dark:text-gray-300">Loading dashboard...</p>
        </div>
      </div>
    )
  }

  if (!userData) {
    return null
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50 dark:bg-gradient-to-br dark:from-blue-900 dark:via-gray-900 dark:to-black">
      {/* Navigation */}
      <nav className="sticky top-0 z-40 bg-white/80 dark:bg-gray-900/80 backdrop-blur-md border-b border-gray-200 dark:border-gray-700">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            {/* Logo */}
            <div className="flex items-center space-x-2">
              <Calendar className="h-6 w-6 sm:h-8 sm:w-8 text-blue-600 dark:text-blue-400" />
              <span className="text-lg sm:text-xl font-bold text-gray-900 dark:text-white">
                <span className="hidden sm:inline">Smart Scheduler</span>
                <span className="sm:hidden">Smart</span>
              </span>
            </div>

            {/* User Info & Actions */}
            <div className="flex items-center space-x-4">
              <div className="hidden md:flex items-center space-x-3">
                <div className="text-right">
                  <p className="text-sm font-medium text-gray-900 dark:text-white">
                    Hello, {userData?.username || "User"}
                  </p>
                  <div className="flex items-center space-x-1">
                    {getRoleIcon(userData?.role || "user")}
                    <Badge className={`text-xs ${getRoleColor(userData?.role || "user")}`}>
                      {userData?.role ? userData.role.charAt(0).toUpperCase() + userData.role.slice(1) : "User"}
                    </Badge>
                  </div>
                </div>
              </div>
              <ThemeToggle />
              <Button
                onClick={handleLogout}
                variant="ghost"
                size="sm"
                className="text-red-600 hover:text-red-700 hover:bg-red-50 dark:text-red-400 dark:hover:text-red-300 dark:hover:bg-red-900/20"
              >
                <LogOut className="h-4 w-4 mr-2" />
                <span className="hidden sm:inline">Logout</span>
              </Button>
            </div>
          </div>
        </div>
      </nav>

      {/* Main Content */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Welcome Section */}
        <div className="mb-8">
          <div className="md:hidden mb-4">
            <div className="text-center p-4 bg-white/80 dark:bg-gray-800/80 backdrop-blur-sm rounded-lg border border-gray-200/50 dark:border-gray-700/50">
              <p className="text-lg font-semibold text-gray-900 dark:text-white mb-2">
                Hello, {userData?.username || "User"}!
              </p>
              <div className="flex items-center justify-center space-x-2">
                {getRoleIcon(userData?.role || "user")}
                <Badge className={`${getRoleColor(userData?.role || "user")}`}>
                  {userData?.role ? userData.role.charAt(0).toUpperCase() + userData.role.slice(1) : "User"}
                </Badge>
              </div>
            </div>
          </div>

          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
            <div>
              <h1 className="text-2xl sm:text-3xl font-bold text-gray-900 dark:text-white">
                {userData?.role ? userData.role.charAt(0).toUpperCase() + userData.role.slice(1) : "User"} Dashboard
              </h1>
              <p className="text-gray-600 dark:text-gray-300 mt-1">Welcome back! Here's your current timetable.</p>
            </div>
            <Button
              onClick={handleDownloadTimetable}
              disabled={isDownloading || !latestFileName}
              className="bg-gradient-to-r from-green-600 to-emerald-600 hover:from-green-700 hover:to-emerald-700 dark:from-green-500 dark:to-emerald-500 dark:hover:from-green-600 dark:hover:to-emerald-600 shadow-lg hover:shadow-xl transition-all duration-200"
            >
              {isDownloading ? (
                <div className="flex items-center">
                  <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                  Downloading...
                </div>
              ) : (
                <div className="flex items-center">
                  <Download className="h-4 w-4 mr-2" />
                  Download Full Timetable
                </div>
              )}
            </Button>
          </div>
        </div>

        {/* Timetable Section */}
        <Card className="shadow-xl border-0 bg-white/90 dark:bg-gray-800/90 backdrop-blur-sm">
          <CardHeader>
            <CardTitle className="flex items-center space-x-2">
              <Calendar className="h-5 w-5 text-blue-600 dark:text-blue-400" />
              <span>Current Timetable</span>
            </CardTitle>
            <CardDescription>
              {latestFileName && (
                <span className="text-sm text-gray-600 dark:text-gray-400">File: {latestFileName}</span>
              )}
            </CardDescription>
          </CardHeader>
          <CardContent>
            {timetableData ? (
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      {timetableData.columns.map((column, index) => (
                        <TableHead key={index} className="whitespace-nowrap font-semibold">
                          {column}
                        </TableHead>
                      ))}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {timetableData.rows.map((row, rowIndex) => (
                      <TableRow key={rowIndex} className="hover:bg-gray-50 dark:hover:bg-gray-800/50">
                        {row.map((cell, cellIndex) => (
                          <TableCell key={cellIndex} className="whitespace-nowrap">
                            {cellIndex === 2 ? ( // CourseType column
                              <Badge className={`${getCourseTypeColor(cell?.toString() || "")}`}>{cell}</Badge>
                            ) : cellIndex === 6 ? ( // Day column
                              <div className="flex items-center space-x-1">
                                <Clock className="h-3 w-3 text-gray-500" />
                                <span className="font-medium">{cell}</span>
                              </div>
                            ) : cellIndex === 9 ? ( // Room column
                              <div className="flex items-center space-x-1">
                                <MapPin className="h-3 w-3 text-gray-500" />
                                <span>{cell}</span>
                              </div>
                            ) : cellIndex === 11 ? ( // Rating column
                              <div className="flex items-center space-x-1">
                                <Star className="h-3 w-3 text-yellow-500 fill-current" />
                                <span>{cell}</span>
                              </div>
                            ) : (
                              cell
                            )}
                          </TableCell>
                        ))}
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            ) : (
              <div className="text-center py-8">
                <Calendar className="h-12 w-12 text-gray-400 mx-auto mb-4" />
                <p className="text-gray-600 dark:text-gray-400">No timetable data available</p>
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    </div>
  )
}
