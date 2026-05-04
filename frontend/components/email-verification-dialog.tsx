"use client"

import type React from "react"

import { useState, useEffect } from "react"
import { Button } from "@/components/ui/button"
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle } from "@/components/ui/dialog"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Mail, RefreshCw, CheckCircle } from "lucide-react"
import { useToast } from "@/hooks/use-toast"
import { API_BASE_URL } from "@/lib/config"

interface EmailVerificationDialogProps {
  isOpen: boolean
  onClose: () => void
  email: string
  onVerificationSuccess: () => void
}

export function EmailVerificationDialog({
  isOpen,
  onClose,
  email,
  onVerificationSuccess,
}: EmailVerificationDialogProps) {
  const [code, setCode] = useState("")
  const [isVerifying, setIsVerifying] = useState(false)
  const [isResending, setIsResending] = useState(false)
  const [resendTimer, setResendTimer] = useState(0)
  const [isVerified, setIsVerified] = useState(false)
  const { toast } = useToast()

  // Start resend timer
  useEffect(() => {
    if (isOpen && !isVerified) {
      setResendTimer(60)
      const timer = setInterval(() => {
        setResendTimer((prev) => {
          if (prev <= 1) {
            clearInterval(timer)
            return 0
          }
          return prev - 1
        })
      }, 1000)

      return () => clearInterval(timer)
    }
  }, [isOpen, isVerified])

  const handleVerifyEmail = async () => {
    if (!code.trim()) {
      toast({
        variant: "destructive",
        title: "Error",
        description: "Please enter the verification code",
      })
      return
    }

    setIsVerifying(true)

    try {
      const response = await fetch(`${API_BASE_URL}/verify-email/`, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email: email,
          code: code.trim(),
        }),
      })

      const data = await response.json()

      if (response.ok) {
        setIsVerified(true)
        toast({
          variant: "success",
          title: "Email Verified!",
          description: data.detail || "Your email has been verified successfully.",
        })
      } else {
        toast({
          variant: "destructive",
          title: "Verification Failed",
          description: data.detail || "Invalid verification code. Please try again.",
        })
      }
    } catch (error) {
      toast({
        variant: "destructive",
        title: "Error",
        description: "Network error. Please check your connection and try again.",
      })
    } finally {
      setIsVerifying(false)
    }
  }

  const handleResendCode = async () => {
    setIsResending(true)

    try {
      const response = await fetch(`${API_BASE_URL}/resend-code/`, {
        method: "POST",
        headers: {
          accept: "application/json",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          email: email,
        }),
      })

      const data = await response.json()

      if (response.ok) {
        toast({
          variant: "success",
          title: "Code Sent!",
          description: data.detail || "A new verification code has been sent to your email.",
        })
        // Restart the timer
        setResendTimer(60)
        const timer = setInterval(() => {
          setResendTimer((prev) => {
            if (prev <= 1) {
              clearInterval(timer)
              return 0
            }
            return prev - 1
          })
        }, 1000)
      } else {
        toast({
          variant: "destructive",
          title: "Resend Failed",
          description: data.detail || "Failed to resend verification code. Please try again.",
        })
      }
    } catch (error) {
      toast({
        variant: "destructive",
        title: "Error",
        description: "Network error. Please check your connection and try again.",
      })
    } finally {
      setIsResending(false)
    }
  }

  const handleCodeChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value.replace(/\D/g, "").slice(0, 4)
    setCode(value)
  }

  const handleClose = () => {
    setIsVerified(false)
    setCode("")
    onClose()
    if (isVerified) {
      onVerificationSuccess()
    }
  }

  return (
    <Dialog open={isOpen} onOpenChange={handleClose}>
      <DialogContent className="sm:max-w-md">
        {!isVerified ? (
          <>
            <DialogHeader className="text-center">
              <div className="flex justify-center mb-4">
                <div className="w-16 h-16 bg-blue-100 dark:bg-blue-900/20 rounded-full flex items-center justify-center">
                  <Mail className="w-8 h-8 text-blue-600 dark:text-blue-400" />
                </div>
              </div>
              <DialogTitle className="text-2xl font-bold">Verify Your Email</DialogTitle>
              <DialogDescription className="text-base">
                We've sent a 4-digit verification code to{" "}
                <span className="font-semibold text-blue-600 dark:text-blue-400">{email}</span>
              </DialogDescription>
            </DialogHeader>

            <div className="space-y-6 py-4">
              <div className="space-y-2">
                <Label htmlFor="code" className="text-sm font-medium">
                  Verification Code
                </Label>
                <Input
                  id="code"
                  type="text"
                  placeholder="Enter 4-digit code"
                  value={code}
                  onChange={handleCodeChange}
                  className="text-center text-2xl font-mono tracking-widest h-14"
                  maxLength={4}
                />
              </div>

              <div className="space-y-3">
                <Button
                  onClick={handleVerifyEmail}
                  disabled={isVerifying || code.length !== 4}
                  className="w-full h-12 bg-blue-600 hover:bg-blue-700 dark:bg-blue-500 dark:hover:bg-blue-600"
                >
                  {isVerifying ? (
                    <div className="flex items-center">
                      <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                      Verifying...
                    </div>
                  ) : (
                    "Verify Email"
                  )}
                </Button>

                <Button
                  onClick={handleResendCode}
                  disabled={isResending || resendTimer > 0}
                  variant="outline"
                  className="w-full h-12"
                >
                  {isResending ? (
                    <div className="flex items-center">
                      <RefreshCw className="h-4 w-4 animate-spin mr-2" />
                      Resending...
                    </div>
                  ) : resendTimer > 0 ? (
                    `Resend Code (${resendTimer}s)`
                  ) : (
                    <div className="flex items-center">
                      <RefreshCw className="h-4 w-4 mr-2" />
                      Resend Code
                    </div>
                  )}
                </Button>
              </div>

              <div className="text-center text-sm text-gray-600 dark:text-gray-400">
                <p>Didn't receive the code? Check your spam folder or try resending.</p>
              </div>
            </div>
          </>
        ) : (
          <>
            <DialogHeader className="text-center">
              <div className="flex justify-center mb-4">
                <div className="w-16 h-16 bg-green-100 dark:bg-green-900/20 rounded-full flex items-center justify-center">
                  <CheckCircle className="w-8 h-8 text-green-600 dark:text-green-400" />
                </div>
              </div>
              <DialogTitle className="text-2xl font-bold text-green-600 dark:text-green-400">
                Email Verified!
              </DialogTitle>
              <DialogDescription className="text-base">Your email has been verified successfully.</DialogDescription>
            </DialogHeader>

            <div className="space-y-6 py-4">
              <div className="text-center p-6 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-200 dark:border-blue-800">
                <h3 className="font-semibold text-blue-900 dark:text-blue-100 mb-2">Registration Pending</h3>
                <p className="text-sm text-blue-700 dark:text-blue-300 leading-relaxed">
                  Your registration request is pending approval. You will receive an email notification once your
                  account is approved or if any additional information is required.
                </p>
              </div>

              <Button onClick={handleClose} className="w-full h-12 bg-blue-600 hover:bg-blue-700">
                Continue
              </Button>
            </div>
          </>
        )}
      </DialogContent>
    </Dialog>
  )
}
