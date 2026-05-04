"use client"

import type React from "react"

import { useEffect, useState } from "react"
import { useRouter } from "next/navigation"
import { API_BASE_URL } from "@/lib/config"
import {
  Calendar,
  LogOut,
  BarChart3,
  History,
  FileText,
  Sparkles,
  Menu,
  X,
  Check,
  XIcon,
  Download,
  Trash2,
  Edit,
  Save,
  RotateCcw,
  Users,
  UserCheck,
  UserX,
  Clock,
} from "lucide-react"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { ThemeToggle } from "@/components/theme-toggle"
import { useToast } from "@/hooks/use-toast"

interface UserData {
  email: string
  username: string
  role: string
}

interface PendingRequest {
  email: string
  username: string
  role: string
}

interface HistoryItem {
  email: string
  username: string
  role: string
  action: string
}

interface TimetableItem {
  filename: string
  modified_time: number
}

interface EditingRow {
  email: string
  role: string
  action: string
}

type ActiveTab = "dashboard" | "history" | "timetable" | "generated-timetable"

export default function AdminDashboardPage() {
  const [userData, setUserData] = useState<UserData | null>(null)
  const [activeTab, setActiveTab] = useState<ActiveTab>("dashboard")
  const [isLoading, setIsLoading] = useState(true)
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false)
  const [mounted, setMounted] = useState(false)

  // Data states
  const [pendingRequests, setPendingRequests] = useState<PendingRequest[]>([])
  const [historyData, setHistoryData] = useState<HistoryItem[]>([])
  const [timetables, setTimetables] = useState<TimetableItem[]>([])
  const [loadingStates, setLoadingStates] = useState<{ [key: string]: boolean }>({})

  // Editing state for history table
  const [editingRow, setEditingRow] = useState<string | null>(null)
  const [editingData, setEditingData] = useState<EditingRow | null>(null)

  // Add these state variables after the existing ones
  const [selectedFile, setSelectedFile] = useState<File | null>(null)
  const [uploadState, setUploadState] = useState({
    isUploading: false,
    uploadedFilename: "",
    isValidating: false,
    validationLog: "",
    isValidated: false,
    isGenerating: false,
    generatedFile: "",
    isDownloading: false,
  })

  const [apiMessage, setApiMessage] = useState("")

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

      if (!isLoggedIn || isLoggedIn !== "true" || !user) {
        router.push("/login")
        return
      }

      const parsedUser = JSON.parse(user)
      setUserData(parsedUser)

      // Load initial data
      loadTabData("dashboard")
    } catch (error) {
      console.error("Error parsing user data:", error)
      router.push("/login")
    } finally {
      setIsLoading(false)
    }
  }, [mounted, router])

  // Load data when tab changes
  useEffect(() => {
    if (userData && mounted) {
      loadTabData(activeTab)
    }
  }, [activeTab, userData, mounted])

  const loadTabData = async (tab: ActiveTab) => {
    if (!userData) return

    try {
      switch (tab) {
        case "dashboard":
          await Promise.all([fetchPendingRequests(), fetchHistoryData()])
          break
        case "history":
          await fetchHistoryData()
          break
        case "generated-timetable":
          await fetchTimetables()
          break
      }
    } catch (error) {
      console.error("Error loading tab data:", error)
    }
  }

  const fetchPendingRequests = async () => {
    try {
      const response = await fetch(`${API_BASE_URL}/admin/requests/`)
      if (!response.ok) throw new Error("Failed to fetch pending requests")

      const data = await response.json()
      setPendingRequests(data.requests || [])
    } catch (error) {
      console.error("Error fetching pending requests:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Failed to load pending requests",
      })
    }
  }

  const fetchHistoryData = async () => {
    try {
      console.log("Fetching history data...")

      const response = await fetch(`${API_BASE_URL}/admin/history/`, {
        method: "GET",
        headers: {
          accept: "application/json",
        },
      })

      console.log("History API response status:", response.status)

      if (!response.ok) throw new Error("Failed to fetch history data")

      const data = await response.json()
      console.log("History API response data:", data)

      setHistoryData(data.history || [])
    } catch (error) {
      console.error("Error fetching history data:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Failed to load history data",
      })
    }
  }

  const fetchTimetables = async () => {
    try {
      const response = await fetch(`${API_BASE_URL}/admin/timetables/`)
      if (!response.ok) throw new Error("Failed to fetch timetables")

      const data = await response.json()
      setTimetables(data.timetables || [])
    } catch (error) {
      console.error("Error fetching timetables:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Failed to load timetables",
      })
    }
  }

  const handleApproveUser = async (email: string) => {
    setLoadingStates((prev) => ({ ...prev, [`approve_${email}`]: true }))

    try {
      const response = await fetch(`${API_BASE_URL}/admin/approve/`, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ email }),
      })

      const data = await response.json()

      if (response.ok) {
        toast({
          variant: "success",
          title: "User Approved",
          description: data.message || "User has been approved successfully",
        })
        // Refresh data
        await Promise.all([fetchPendingRequests(), fetchHistoryData()])
      } else {
        toast({
          variant: "destructive",
          title: "Approval Failed",
          description: data.detail || data.message || "Failed to approve user",
        })
      }
    } catch (error) {
      console.error("Error approving user:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Network error occurred",
      })
    } finally {
      setLoadingStates((prev) => ({ ...prev, [`approve_${email}`]: false }))
    }
  }

  const handleRejectUser = async (email: string) => {
    setLoadingStates((prev) => ({ ...prev, [`reject_${email}`]: true }))

    try {
      const response = await fetch(`${API_BASE_URL}/admin/reject/`, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ email }),
      })

      const data = await response.json()

      if (response.ok) {
        toast({
          variant: "success",
          title: "User Rejected",
          description: data.message || "User has been rejected successfully",
        })
        // Refresh data
        await Promise.all([fetchPendingRequests(), fetchHistoryData()])
      } else {
        toast({
          variant: "destructive",
          title: "Rejection Failed",
          description: data.detail || data.message || "Failed to reject user",
        })
      }
    } catch (error) {
      console.error("Error rejecting user:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Network error occurred",
      })
    } finally {
      setLoadingStates((prev) => ({ ...prev, [`reject_${email}`]: false }))
    }
  }

  const handleEditHistory = (item: HistoryItem) => {
    setEditingRow(item.email)
    setEditingData({
      email: item.email,
      role: item.role,
      action: item.action,
    })
  }

  const handleSaveEdit = async () => {
    if (!editingData) return

    console.log("Saving edit for:", editingData)
    setLoadingStates((prev) => ({ ...prev, [`save_${editingData.email}`]: true }))

    try {
      const response = await fetch(`${API_BASE_URL}/admin/history/edit/`, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email: editingData.email,
          action: editingData.action,
          role: editingData.role,
        }),
      })

      const data = await response.json()
      console.log("Edit API response:", data)

      if (response.ok) {
        toast({
          variant: "success",
          title: "History Updated",
          description: data.message || "History record updated successfully",
        })
        setEditingRow(null)
        setEditingData(null)
        await fetchHistoryData()
      } else {
        toast({
          variant: "destructive",
          title: "Update Failed",
          description: data.detail || data.message || "Failed to update history record",
        })
      }
    } catch (error) {
      console.error("Error updating history:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Network error occurred",
      })
    } finally {
      setLoadingStates((prev) => ({ ...prev, [`save_${editingData.email}`]: false }))
    }
  }

  const handleCancelEdit = () => {
    setEditingRow(null)
    setEditingData(null)
  }

  const handleDeleteHistory = async (email: string) => {
    console.log("Deleting history for email:", email)
    setLoadingStates((prev) => ({ ...prev, [`delete_${email}`]: true }))

    try {
      const response = await fetch(`${API_BASE_URL}/admin/history/remove/`, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ email }),
      })

      const data = await response.json()
      console.log("Delete API response:", data)

      if (response.ok) {
        toast({
          variant: "success",
          title: "Record Deleted",
          description: data.message || "History record deleted successfully",
        })
        await fetchHistoryData()
      } else {
        toast({
          variant: "destructive",
          title: "Delete Failed",
          description: data.detail || data.message || "Failed to delete history record",
        })
      }
    } catch (error) {
      console.error("Error deleting history:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Network error occurred",
      })
    } finally {
      setLoadingStates((prev) => ({ ...prev, [`delete_${email}`]: false }))
    }
  }

  const handleDownloadTimetable = async (filename: string) => {
    setLoadingStates((prev) => ({ ...prev, [`download_${filename}`]: true }))

    try {
      const response = await fetch(`${API_BASE_URL}/download/${filename}`)

      if (!response.ok) throw new Error("Failed to download timetable")

      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement("a")
      a.style.display = "none"
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)

      toast({
        variant: "success",
        title: "Download Complete",
        description: "Timetable downloaded successfully",
      })
    } catch (error) {
      console.error("Error downloading timetable:", error)
      toast({
        variant: "destructive",
        title: "Download Failed",
        description: "Failed to download timetable",
      })
    } finally {
      setLoadingStates((prev) => ({ ...prev, [`download_${filename}`]: false }))
    }
  }

  const handleDeleteTimetable = async (filename: string) => {
    setLoadingStates((prev) => ({ ...prev, [`delete_timetable_${filename}`]: true }))

    try {
      const response = await fetch(`${API_BASE_URL}/admin/timetables/${filename}`, {
        method: "DELETE",
      })

      const data = await response.json()

      if (response.ok) {
        toast({
          variant: "success",
          title: "Timetable Deleted",
          description: data.message || "Timetable deleted successfully",
        })
        await fetchTimetables()
      } else {
        toast({
          variant: "destructive",
          title: "Delete Failed",
          description: data.detail || data.message || "Failed to delete timetable",
        })
      }
    } catch (error) {
      console.error("Error deleting timetable:", error)
      toast({
        variant: "destructive",
        title: "Error",
        description: "Network error occurred",
      })
    } finally {
      setLoadingStates((prev) => ({ ...prev, [`delete_timetable_${filename}`]: false }))
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

  const tabs = [
    { id: "dashboard" as ActiveTab, label: "Dashboard", icon: BarChart3 },
    { id: "history" as ActiveTab, label: "History", icon: History },
    { id: "timetable" as ActiveTab, label: "Timetable", icon: FileText },
    { id: "generated-timetable" as ActiveTab, label: "Generated Timetable", icon: Sparkles },
  ]

  // Calculate statistics
  const totalUsers = historyData.length
  const pendingRequestsCount = pendingRequests.length
  const approvedUsers = historyData.filter((item) => item.action === "approved").length
  const rejectedUsers = historyData.filter((item) => item.action === "rejected").length

  // Add these handler functions after the existing ones
  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (file) {
      // Check if file is Excel format
      const validTypes = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
      ]

      if (validTypes.includes(file.type) || file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
        setSelectedFile(file)
        // Reset upload state when new file is selected
        setUploadState({
          isUploading: false,
          uploadedFilename: "",
          isValidating: false,
          validationLog: "",
          isValidated: false,
          isGenerating: false,
          generatedFile: "",
          isDownloading: false,
        })
        setApiMessage("")
      } else {
        toast({
          variant: "destructive",
          title: "Invalid File Type",
          description: "Please select an Excel file (.xlsx or .xls)",
        })
      }
    }
  }

  const handleUploadFile = async () => {
    if (!selectedFile) return

    console.log("Starting file upload:", selectedFile.name)
    setUploadState((prev) => ({ ...prev, isUploading: true }))

    try {
      const formData = new FormData()
      formData.append("excel_file", selectedFile)

        const response = await fetch(`${API_BASE_URL}/upload-excel/`, {
        method: "POST",
        body: formData,
      })

      const data = await response.json()
      console.log("Upload API response:", {
        status: response.status,
        data: data,
      })

      if (response.ok) {
        toast({
          variant: "success",
          title: "Upload Successful",
          description: data.message || "Excel file uploaded successfully",
        })
        setUploadState((prev) => ({
          ...prev,
          isUploading: false,
          uploadedFilename: data.filename || selectedFile.name,
        }))
      } else {
        console.error("Upload failed:", data)
        toast({
          variant: "destructive",
          title: "Upload Failed",
          description: data.detail || data.message || "Failed to upload Excel file",
        })
        setUploadState((prev) => ({ ...prev, isUploading: false }))
      }
    } catch (error) {
      console.error("Upload error:", error)
      toast({
        variant: "destructive",
        title: "Upload Error",
        description: "Network error occurred during upload",
      })
      setUploadState((prev) => ({ ...prev, isUploading: false }))
    }
  }

  const handleValidateFile = async () => {
    if (!uploadState.uploadedFilename) return

    console.log("Starting file validation:", uploadState.uploadedFilename)
    setUploadState((prev) => ({ ...prev, isValidating: true, validationLog: "" }))

    try {
      const formData = new URLSearchParams()
      formData.append("filename", uploadState.uploadedFilename)

      const response = await fetch(`${API_BASE_URL}/validate-excel/`, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: formData,
      })

      const data = await response.json()
      console.log("Validation API response:", {
        status: response.status,
        data: data,
      })

      if (response.ok) {
        const message = data.message || "Excel data validated successfully"
        setApiMessage(message)
        toast({
          variant: "success",
          title: "Validation Successful",
          description: message,
        })
        setUploadState((prev) => ({
          ...prev,
          isValidating: false,
          validationLog: data.validation_log || "",
          isValidated: true,
        }))
      } else {
        console.error("Validation failed:", data)
        toast({
          variant: "destructive",
          title: "Validation Failed",
          description: data.detail || data.message || "Failed to validate Excel file",
        })
        setUploadState((prev) => ({ ...prev, isValidating: false }))
      }
    } catch (error) {
      console.error("Validation error:", error)
      toast({
        variant: "destructive",
        title: "Validation Error",
        description: "Network error occurred during validation",
      })
      setUploadState((prev) => ({ ...prev, isValidating: false }))
    }
  }

  const handleGenerateTimetable = async () => {
    if (!uploadState.uploadedFilename) return

    console.log("Starting timetable generation:", uploadState.uploadedFilename)
    setUploadState((prev) => ({ ...prev, isGenerating: true }))

    try {
      const formData = new URLSearchParams()
      formData.append("filename", uploadState.uploadedFilename)

      const response = await fetch(`${API_BASE_URL}/generate-timetable/`, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: formData,
      })

      const data = await response.json()
      console.log("Generation API response:", {
        status: response.status,
        data: data,
      })

      if (response.ok) {
        toast({
          variant: "success",
          title: "Generation Successful",
          description: data.message || "Timetable generated successfully",
        })
        setUploadState((prev) => ({
          ...prev,
          isGenerating: false,
          generatedFile: data.output_file || "",
        }))
        // Refresh the generated timetables list
        await fetchTimetables()
      } else {
        console.error("Generation failed:", data)
        toast({
          variant: "destructive",
          title: "Generation Failed",
          description: data.detail || data.message || "Failed to generate timetable",
        })
        setUploadState((prev) => ({ ...prev, isGenerating: false }))
      }
    } catch (error) {
      console.error("Generation error:", error)
      toast({
        variant: "destructive",
        title: "Generation Error",
        description: "Network error occurred during timetable generation",
      })
      setUploadState((prev) => ({ ...prev, isGenerating: false }))
    }
  }

  const handleDownloadGenerated = async () => {
    if (!uploadState.generatedFile) return

    console.log("Starting download:", uploadState.generatedFile)
    setUploadState((prev) => ({ ...prev, isDownloading: true }))

    try {
      const response = await fetch(`${API_BASE_URL}/download/${uploadState.generatedFile}`)

      console.log("Download API response:", {
        status: response.status,
        headers: Object.fromEntries(response.headers.entries()),
      })

      if (!response.ok) throw new Error("Failed to download timetable")

      const blob = await response.blob()
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement("a")
      a.style.display = "none"
      a.href = url
      a.download = uploadState.generatedFile
      document.body.appendChild(a)
      a.click()
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)

      console.log("Download completed successfully")
      toast({
        variant: "success",
        title: "Download Complete",
        description: "Timetable downloaded successfully",
      })
    } catch (error) {
      console.error("Download error:", error)
      toast({
        variant: "destructive",
        title: "Download Failed",
        description: "Failed to download generated timetable",
      })
    } finally {
      setUploadState((prev) => ({ ...prev, isDownloading: false }))
    }
  }

  const renderTabContent = () => {
    switch (activeTab) {
      case "dashboard":
        return (
          <div className="space-y-6">
            {/* Statistics Cards */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
              <Card className="bg-gradient-to-r from-blue-500 to-blue-600 text-white border-0">
                <CardContent className="p-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-blue-100">Total Users</p>
                      <p className="text-3xl font-bold">{totalUsers}</p>
                    </div>
                    <Users className="h-8 w-8 text-blue-200" />
                  </div>
                </CardContent>
              </Card>
              <Card className="bg-gradient-to-r from-yellow-500 to-yellow-600 text-white border-0">
                <CardContent className="p-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-yellow-100">Pending Requests</p>
                      <p className="text-3xl font-bold">{pendingRequestsCount}</p>
                    </div>
                    <Clock className="h-8 w-8 text-yellow-200" />
                  </div>
                </CardContent>
              </Card>
              <Card className="bg-gradient-to-r from-green-500 to-green-600 text-white border-0">
                <CardContent className="p-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-green-100">Accepted Users</p>
                      <p className="text-3xl font-bold">{approvedUsers}</p>
                    </div>
                    <UserCheck className="h-8 w-8 text-green-200" />
                  </div>
                </CardContent>
              </Card>
              <Card className="bg-gradient-to-r from-red-500 to-red-600 text-white border-0">
                <CardContent className="p-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-red-100">Rejected Users</p>
                      <p className="text-3xl font-bold">{rejectedUsers}</p>
                    </div>
                    <UserX className="h-8 w-8 text-red-200" />
                  </div>
                </CardContent>
              </Card>
            </div>

            {/* Pending Requests Table */}
            <Card>
              <CardHeader>
                <CardTitle>Pending User Requests</CardTitle>
                <CardDescription>Review and approve/reject user registration requests</CardDescription>
              </CardHeader>
              <CardContent>
                {pendingRequests.length > 0 ? (
                  <div className="overflow-x-auto">
                    <Table>
                      <TableHeader>
                        <TableRow>
                          <TableHead>Name</TableHead>
                          <TableHead>Email</TableHead>
                          <TableHead>Role</TableHead>
                          <TableHead>Actions</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {pendingRequests.map((request) => (
                          <TableRow key={request.email}>
                            <TableCell className="font-medium">{request.username}</TableCell>
                            <TableCell>{request.email}</TableCell>
                            <TableCell>
                              <span className="capitalize bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200 px-2 py-1 rounded-full text-xs font-medium">
                                {request.role}
                              </span>
                            </TableCell>
                            <TableCell>
                              <div className="flex space-x-2">
                                <Button
                                  size="sm"
                                  onClick={() => handleApproveUser(request.email)}
                                  disabled={loadingStates[`approve_${request.email}`]}
                                  className="bg-green-600 hover:bg-green-700 text-white"
                                >
                                  {loadingStates[`approve_${request.email}`] ? (
                                    <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-white" />
                                  ) : (
                                    <Check className="h-3 w-3" />
                                  )}
                                  <span className="ml-1 hidden sm:inline">Approve</span>
                                </Button>
                                <Button
                                  size="sm"
                                  variant="destructive"
                                  onClick={() => handleRejectUser(request.email)}
                                  disabled={loadingStates[`reject_${request.email}`]}
                                >
                                  {loadingStates[`reject_${request.email}`] ? (
                                    <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-white" />
                                  ) : (
                                    <XIcon className="h-3 w-3" />
                                  )}
                                  <span className="ml-1 hidden sm:inline">Reject</span>
                                </Button>
                              </div>
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                ) : (
                  <div className="text-center py-8">
                    <Clock className="h-12 w-12 text-gray-400 mx-auto mb-4" />
                    <p className="text-gray-600 dark:text-gray-400">No pending requests</p>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>
        )

      case "history":
        return (
          <Card>
            <CardHeader>
              <CardTitle>User History</CardTitle>
              <CardDescription>View and manage user approval history</CardDescription>
            </CardHeader>
            <CardContent>
              {historyData.length > 0 ? (
                <div className="overflow-x-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Name</TableHead>
                        <TableHead>Email</TableHead>
                        <TableHead>Role</TableHead>
                        <TableHead>Status</TableHead>
                        <TableHead>Actions</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {historyData.map((item) => (
                        <TableRow key={item.email}>
                          <TableCell className="font-medium">{item.username}</TableCell>
                          <TableCell>{item.email}</TableCell>
                          <TableCell>
                            {editingRow === item.email ? (
                              <Select
                                value={editingData?.role}
                                onValueChange={(value) =>
                                  setEditingData((prev) => (prev ? { ...prev, role: value } : null))
                                }
                              >
                                <SelectTrigger className="w-32">
                                  <SelectValue />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="student">Student</SelectItem>
                                  <SelectItem value="teacher">Teacher</SelectItem>
                                  <SelectItem value="admin">Admin</SelectItem>
                                </SelectContent>
                              </Select>
                            ) : (
                              <span className="capitalize bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200 px-2 py-1 rounded-full text-xs font-medium">
                                {item.role}
                              </span>
                            )}
                          </TableCell>
                          <TableCell>
                            {editingRow === item.email ? (
                              <Select
                                value={editingData?.action}
                                onValueChange={(value) =>
                                  setEditingData((prev) => (prev ? { ...prev, action: value } : null))
                                }
                              >
                                <SelectTrigger className="w-32">
                                  <SelectValue />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="approved">Approved</SelectItem>
                                  <SelectItem value="rejected">Rejected</SelectItem>
                                </SelectContent>
                              </Select>
                            ) : (
                              <span
                                className={`capitalize px-2 py-1 rounded-full text-xs font-medium ${
                                  item.action === "approved"
                                    ? "bg-green-100 text-green-800 dark:bg-green-900 dark:text-green-200"
                                    : item.action === "rejected"
                                      ? "bg-red-100 text-red-800 dark:bg-red-900 dark:text-red-200"
                                      : "bg-yellow-100 text-yellow-800 dark:bg-yellow-900 dark:text-yellow-200"
                                }`}
                              >
                                {item.action}
                              </span>
                            )}
                          </TableCell>
                          <TableCell>
                            {editingRow === item.email ? (
                              <div className="flex space-x-2">
                                <Button
                                  size="sm"
                                  onClick={handleSaveEdit}
                                  disabled={loadingStates[`save_${item.email}`]}
                                  className="bg-green-600 hover:bg-green-700 text-white"
                                >
                                  {loadingStates[`save_${item.email}`] ? (
                                    <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-white" />
                                  ) : (
                                    <Save className="h-3 w-3" />
                                  )}
                                </Button>
                                <Button size="sm" variant="outline" onClick={handleCancelEdit}>
                                  <RotateCcw className="h-3 w-3" />
                                </Button>
                              </div>
                            ) : (
                              <div className="flex space-x-2">
                                <Button size="sm" variant="outline" onClick={() => handleEditHistory(item)}>
                                  <Edit className="h-3 w-3" />
                                </Button>
                                <Button
                                  size="sm"
                                  variant="destructive"
                                  onClick={() => handleDeleteHistory(item.email)}
                                  disabled={loadingStates[`delete_${item.email}`]}
                                >
                                  {loadingStates[`delete_${item.email}`] ? (
                                    <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-white" />
                                  ) : (
                                    <Trash2 className="h-3 w-3" />
                                  )}
                                </Button>
                              </div>
                            )}
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              ) : (
                <div className="text-center py-8">
                  <History className="h-12 w-12 text-gray-400 mx-auto mb-4" />
                  <p className="text-gray-600 dark:text-gray-400">No history records found</p>
                </div>
              )}
            </CardContent>
          </Card>
        )

      case "timetable":
        return (
          <Card>
            <CardHeader>
              <CardTitle>Timetable Management</CardTitle>
              <CardDescription>Upload Excel file and generate AI-powered timetables</CardDescription>
            </CardHeader>
            <CardContent className="space-y-8">
              {/* File Upload Area */}
              <div className="border-2 border-dashed border-gray-300 dark:border-gray-600 rounded-lg p-12 text-center">
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileSelect}
                  className="hidden"
                  id="excel-upload"
                />
                <label htmlFor="excel-upload" className="cursor-pointer flex flex-col items-center space-y-6">
                  <FileText className="h-16 w-16 text-gray-400" />
                  {selectedFile ? (
                    <div className="space-y-3">
                      <p className="text-xl font-medium text-gray-900 dark:text-white">{selectedFile.name}</p>
                      <p className="text-base text-gray-500">{(selectedFile.size / 1024 / 1024).toFixed(2)} MB</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      <p className="text-xl font-medium text-gray-900 dark:text-white">Click to upload Excel file</p>
                      <p className="text-base text-gray-500">Only .xlsx and .xls files are supported</p>
                    </div>
                  )}
                </label>
              </div>

              {/* Single Action Button */}
              <div className="flex justify-center">
                <Button
                  onClick={
                    !uploadState.uploadedFilename
                      ? handleUploadFile
                      : !uploadState.isValidated
                        ? handleValidateFile
                        : !uploadState.generatedFile
                          ? handleGenerateTimetable
                          : handleDownloadGenerated
                  }
                  disabled={
                    !selectedFile ||
                    uploadState.isUploading ||
                    uploadState.isValidating ||
                    uploadState.isGenerating ||
                    uploadState.isDownloading
                  }
                  className="px-16 py-4 text-lg font-medium min-w-[300px] bg-blue-600 hover:bg-blue-700 text-white"
                >
                  {uploadState.isUploading ? (
                    <>
                      <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white mr-3" />
                      Uploading...
                    </>
                  ) : uploadState.isValidating ? (
                    <>
                      <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white mr-3" />
                      Validating...
                    </>
                  ) : uploadState.isGenerating ? (
                    <>
                      <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white mr-3" />
                      Generating...
                    </>
                  ) : uploadState.isDownloading ? (
                    <>
                      <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white mr-3" />
                      Downloading...
                    </>
                  ) : !uploadState.uploadedFilename ? (
                    "Upload"
                  ) : !uploadState.isValidated ? (
                    "Validate"
                  ) : !uploadState.generatedFile ? (
                    "Generate Timetable"
                  ) : (
                    <>
                      <Download className="h-5 w-5 mr-3" />
                      Download Timetable
                    </>
                  )}
                </Button>
              </div>
            </CardContent>
          </Card>
        )

      case "generated-timetable":
        return (
          <Card>
            <CardHeader>
              <CardTitle>Generated Timetables</CardTitle>
              <CardDescription>View, download, and manage AI-generated timetables</CardDescription>
            </CardHeader>
            <CardContent>
              {timetables.length > 0 ? (
                <div className="overflow-x-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Timetable</TableHead>
                        <TableHead>Download</TableHead>
                        <TableHead>Delete</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {timetables.map((timetable) => (
                        <TableRow key={timetable.filename}>
                          <TableCell className="font-medium">{timetable.filename}</TableCell>
                          <TableCell>
                            <Button
                              size="sm"
                              onClick={() => handleDownloadTimetable(timetable.filename)}
                              disabled={loadingStates[`download_${timetable.filename}`]}
                              className="bg-blue-600 hover:bg-blue-700 text-white"
                            >
                              {loadingStates[`download_${timetable.filename}`] ? (
                                <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-white" />
                              ) : (
                                <Download className="h-3 w-3" />
                              )}
                              <span className="ml-1 hidden sm:inline">Download</span>
                            </Button>
                          </TableCell>
                          <TableCell>
                            <Button
                              size="sm"
                              variant="destructive"
                              onClick={() => handleDeleteTimetable(timetable.filename)}
                              disabled={loadingStates[`delete_timetable_${timetable.filename}`]}
                            >
                              {loadingStates[`delete_timetable_${timetable.filename}`] ? (
                                <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-white" />
                              ) : (
                                <Trash2 className="h-3 w-3" />
                              )}
                              <span className="ml-1 hidden sm:inline">Delete</span>
                            </Button>
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              ) : (
                <div className="text-center py-8">
                  <Sparkles className="h-12 w-12 text-gray-400 mx-auto mb-4" />
                  <p className="text-gray-600 dark:text-gray-400">No generated timetables found</p>
                </div>
              )}
            </CardContent>
          </Card>
        )

      default:
        return null
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
          <p className="text-gray-600 dark:text-gray-300">Loading admin dashboard...</p>
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
            {/* Logo - Left Side */}
            <div className="flex items-center space-x-2">
              <Calendar className="h-6 w-6 sm:h-8 sm:w-8 text-blue-600 dark:text-blue-400" />
              <span className="text-lg sm:text-xl font-bold text-gray-900 dark:text-white">
                <span className="hidden sm:inline">Smart Scheduler</span>
                <span className="sm:hidden">Smart</span>
              </span>
            </div>

            {/* Desktop Navigation - Right Side */}
            <div className="hidden lg:flex items-center space-x-1">
              {tabs.map((tab) => {
                const Icon = tab.icon
                return (
                  <Button
                    key={tab.id}
                    variant={activeTab === tab.id ? "default" : "ghost"}
                    onClick={() => setActiveTab(tab.id)}
                    className={`flex items-center space-x-2 px-4 py-2 ${
                      activeTab === tab.id
                        ? "bg-blue-600 text-white hover:bg-blue-700 dark:bg-blue-500 dark:hover:bg-blue-600"
                        : "hover:bg-blue-50 dark:hover:bg-blue-900/20"
                    }`}
                  >
                    <Icon className="h-4 w-4" />
                    <span className="hidden xl:inline">{tab.label}</span>
                  </Button>
                )
              })}
              <div className="ml-4 flex items-center space-x-2">
                <ThemeToggle />
                <Button
                  onClick={handleLogout}
                  variant="ghost"
                  size="sm"
                  className="text-red-600 hover:text-red-700 hover:bg-red-50 dark:text-red-400 dark:hover:text-red-300 dark:hover:bg-red-900/20"
                >
                  <LogOut className="h-4 w-4" />
                </Button>
              </div>
            </div>

            {/* Mobile Menu Button */}
            <div className="lg:hidden flex items-center space-x-2">
              <ThemeToggle />
              <Button variant="ghost" size="sm" onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)} className="p-2">
                {isMobileMenuOpen ? <X className="h-5 w-5" /> : <Menu className="h-5 w-5" />}
              </Button>
            </div>
          </div>

          {/* Mobile Navigation Menu */}
          {isMobileMenuOpen && (
            <div className="lg:hidden border-t border-gray-200 dark:border-gray-700 py-4">
              <div className="space-y-2">
                {tabs.map((tab) => {
                  const Icon = tab.icon
                  return (
                    <Button
                      key={tab.id}
                      variant={activeTab === tab.id ? "default" : "ghost"}
                      onClick={() => {
                        setActiveTab(tab.id)
                        setIsMobileMenuOpen(false)
                      }}
                      className={`w-full justify-start flex items-center space-x-2 ${
                        activeTab === tab.id
                          ? "bg-blue-600 text-white hover:bg-blue-700 dark:bg-blue-500 dark:hover:bg-blue-600"
                          : "hover:bg-blue-50 dark:hover:bg-blue-900/20"
                      }`}
                    >
                      <Icon className="h-4 w-4" />
                      <span>{tab.label}</span>
                    </Button>
                  )
                })}
                <div className="pt-2 border-t border-gray-200 dark:border-gray-700">
                  <div className="flex items-center justify-between px-4 py-2">
                    <div className="text-sm">
                      <p className="font-medium text-gray-900 dark:text-white">
                        Admin: {userData?.username || "Admin"}
                      </p>
                      <p className="text-gray-600 dark:text-gray-400">{userData?.email}</p>
                    </div>
                    <Button
                      onClick={handleLogout}
                      variant="ghost"
                      size="sm"
                      className="text-red-600 hover:text-red-700 hover:bg-red-50 dark:text-red-400 dark:hover:text-red-300 dark:hover:bg-red-900/20"
                    >
                      <LogOut className="h-4 w-4" />
                    </Button>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </nav>

      {/* Main Content */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Welcome Section */}
        <div className="mb-8">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
            <div>
              <h1 className="text-2xl sm:text-3xl font-bold text-gray-900 dark:text-white">Admin Dashboard</h1>
              <p className="text-gray-600 dark:text-gray-300 mt-1">
                Welcome back, {userData?.username || "Admin"}! Manage your system from here.
              </p>
            </div>
            <div className="hidden sm:block">
              <div className="text-right">
                <p className="text-sm font-medium text-gray-900 dark:text-white">{userData?.username || "Admin"}</p>
                <p className="text-sm text-gray-600 dark:text-gray-400">{userData?.email}</p>
              </div>
            </div>
          </div>
        </div>

        {/* Tab Content */}
        {renderTabContent()}
      </div>
    </div>
  )
}
