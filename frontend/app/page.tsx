import Link from "next/link"
import { Calendar, Clock, Users, Heart, Upload, Cpu, Download, CheckCircle, XCircle, UserCog, GraduationCap, BookOpen } from "lucide-react"
import { Button } from "@/components/ui/button"
import { Card, CardContent } from "@/components/ui/card"
import { ThemeToggle } from "@/components/theme-toggle"
import { MobileNav } from "@/components/mobile-nav"

export default function HomePage() {
  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-indigo-50 dark:bg-gradient-to-br dark:from-blue-900 dark:via-gray-900 dark:to-black">
      {/* Navigation */}
      <nav className="sticky top-0 z-40 bg-white/80 dark:bg-gray-900/80 backdrop-blur-md border-b border-gray-200 dark:border-gray-700">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            {/* Logo */}
            <div className="flex items-center space-x-2">
              <Calendar className="h-6 w-6 sm:h-8 sm:w-8 text-blue-600 dark:text-blue-400" />
              <span className="text-lg sm:text-xl font-bold text-gray-900 dark:text-white">Smart Scheduler</span>
            </div>

            {/* Desktop Navigation */}
            <div className="hidden md:flex items-center space-x-4">
              <ThemeToggle />
              <Link href="/login">
                <Button variant="ghost" className="hover:bg-blue-50 dark:hover:bg-blue-900/20">
                  Login
                </Button>
              </Link>
              <Link href="/signup">
                <Button className="bg-blue-600 hover:bg-blue-700 dark:bg-blue-500 dark:hover:bg-blue-600">
                  Sign Up
                </Button>
              </Link>
            </div>

            {/* Mobile Navigation */}
            <MobileNav />
          </div>
        </div>
      </nav>

      {/* Hero Section */}
      <section className="py-8 sm:py-12 lg:py-20 px-4 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto">
          <div className="grid lg:grid-cols-2 gap-8 lg:gap-12 items-center">
            {/* Hero Text - Left Side */}
            <div className="order-1 text-center lg:text-left">
              <h1 className="text-3xl sm:text-4xl lg:text-6xl font-bold text-gray-900 dark:text-white mb-4 sm:mb-6 leading-tight">
                <span className="bg-gradient-to-r from-blue-600 via-indigo-600 to-purple-600 dark:from-blue-400 dark:via-indigo-400 dark:to-purple-400 bg-clip-text text-transparent">
                  Smart
                </span>{" "}
                Scheduler
              </h1>
              <p className="text-base sm:text-lg text-gray-600 dark:text-gray-300 mb-6 sm:mb-8 leading-relaxed">
                University scheduling takes days. Ours takes minutes. Stop fighting spreadsheets. Let AI build your semester timetable.
              </p>
              <p className="text-base sm:text-lg text-gray-600 dark:text-gray-300 mb-6 sm:mb-8 leading-relaxed">
                Upload your faculty, courses and rooms data. Smart Scheduler generates a conflict-free timetable — and flags anything it couldn't resolve, so you know exactly where to look.
              </p>
              <div className="flex flex-col sm:flex-row gap-3 sm:gap-4 justify-center lg:justify-start">
                <Link href="/signup" className="w-full sm:w-auto">
                  <Button
                    size="lg"
                    className="w-full sm:w-auto bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 dark:from-blue-500 dark:to-indigo-500 dark:hover:from-blue-600 dark:hover:to-indigo-600 shadow-lg hover:shadow-xl transition-all duration-200"
                  >
                    Get Started Free
                  </Button>
                </Link>
              </div>

              {/* Stats - Hidden on mobile */}
              <div className="hidden sm:grid grid-cols-4 gap-4 lg:gap-6 pt-8 mt-8 border-t border-gray-200/50 dark:border-gray-700/50">
                {[
                  { number: "5 min", label: "Avg Runtime Per Schedule" },
                  { number: "1000+", label: "Students Per Schedule" },
                  { number: "50+", label: "Courses Per Semester" },
                  { number: "Days", label: "Of Scheduling Time Saved" },
                ].map((stat, index) => (
                  <div key={index} className="text-center">
                    <div className="text-xl lg:text-2xl font-bold bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-400 bg-clip-text text-transparent">
                      {stat.number}
                    </div>
                    <div className="text-sm text-gray-600 dark:text-gray-400">{stat.label}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Sample Timetable - Right Side */}
            <div className="order-2">
              <div className="bg-white/90 dark:bg-gray-800/90 backdrop-blur-sm rounded-xl sm:rounded-2xl shadow-lg sm:shadow-xl p-4 sm:p-6 border border-gray-200/50 dark:border-gray-700/50">
                <h3 className="text-base sm:text-lg font-semibold text-gray-900 dark:text-white mb-3 sm:mb-4 flex items-center">
                  <Clock className="h-4 w-4 sm:h-5 sm:w-5 text-blue-600 dark:text-blue-400 mr-2" />
                  Today's Schedule
                </h3>
                <div className="space-y-2 sm:space-y-3">
                  {[
                    { time: "9:00 AM", subject: "Programming Fundamentals", room: "101", color: "bg-gradient-to-r from-blue-500 to-blue-600 text-white" },
                    { time: "10:30 AM", subject: "Physics", room: "Lab 2", color: "bg-gradient-to-r from-green-500 to-green-600 text-white" },
                    { time: "12:00 PM", subject: "Lunch", room: "Cafe", color: "bg-gradient-to-r from-yellow-500 to-orange-500 text-white" },
                    { time: "1:00 PM", subject: "Discrete Mathematics", room: "Lab 1", color: "bg-gradient-to-r from-purple-500 to-purple-600 text-white" },
                    { time: "2:30 PM", subject: "English I", room: "205", color: "bg-gradient-to-r from-pink-500 to-pink-600 text-white" },
                  ].map((item, index) => (
                    <div
                      key={index}
                      className="flex items-center justify-between p-2 sm:p-3 rounded-lg border border-gray-100/50 dark:border-gray-700/50 hover:shadow-md transition-all duration-200 bg-white/50 dark:bg-gray-800/50"
                    >
                      <div className="flex items-center space-x-2 sm:space-x-3 flex-1 min-w-0">
                        <div className="text-xs sm:text-sm font-medium text-gray-600 dark:text-gray-300 w-16 sm:w-20 flex-shrink-0">
                          {item.time}
                        </div>
                        <div className={`px-2 sm:px-3 py-1 rounded-lg text-xs font-medium ${item.color} truncate shadow-sm`}>
                          {item.subject}
                        </div>
                      </div>
                      <div className="text-xs text-gray-500 dark:text-gray-400 flex-shrink-0 ml-2">{item.room}</div>
                    </div>
                  ))}
                </div>
                <div className="mt-4 pt-3 border-t border-gray-100/50 dark:border-gray-700/50">
                  <div className="flex items-center justify-between text-xs sm:text-sm">
                    <span className="text-gray-600 dark:text-gray-300">5 classes today</span>
                    <span className="text-green-500 dark:text-green-400 font-semibold">No conflicts</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* How It Works Section */}
      <section className="py-12 sm:py-16 lg:py-20 px-4 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto">
          <div className="text-center mb-12 sm:mb-16">
            <h2 className="text-2xl sm:text-3xl lg:text-4xl font-bold text-gray-900 dark:text-white mb-3 sm:mb-4">
              How It{" "}
              <span className="bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-400 bg-clip-text text-transparent">
                Works
              </span>
            </h2>
            <p className="text-base sm:text-lg text-gray-600 dark:text-gray-300 max-w-2xl mx-auto">
              Three steps. No training required. No scheduling expertise needed.
            </p>
          </div>

          <div className="grid sm:grid-cols-3 gap-6 sm:gap-8 relative">
            {/* Connector line — desktop only */}
            <div className="hidden sm:block absolute top-10 left-1/3 right-1/3 h-0.5 bg-gradient-to-r from-blue-300 to-indigo-300 dark:from-blue-700 dark:to-indigo-700 z-0" />

            {[
              {
                step: "01",
                icon: Upload,
                title: "Upload Your Data",
                description: "One Excel file with your faculty, courses, rooms, and time slots. That's all the system needs to get started.",
                iconColor: "text-blue-600 dark:text-blue-400",
                bgColor: "bg-blue-50 dark:bg-blue-900/20",
              },
              {
                step: "02",
                icon: Cpu,
                title: "AI Runs the Schedule",
                description: "The system applies Constraint Satisfaction and Genetic Algorithms to assign every course, room, and faculty slot — conflict-free.",
                iconColor: "text-indigo-600 dark:text-indigo-400",
                bgColor: "bg-indigo-50 dark:bg-indigo-900/20",
              },
              {
                step: "03",
                icon: Download,
                title: "Download & Share",
                description: "Your complete timetable is ready to download as an Excel file. Faculty and students can view their schedules instantly.",
                iconColor: "text-purple-600 dark:text-purple-400",
                bgColor: "bg-purple-50 dark:bg-purple-900/20",
              },
            ].map((step, index) => (
              <div key={index} className="relative z-10 text-center">
                <div className={`inline-flex items-center justify-center w-16 h-16 sm:w-20 sm:h-20 rounded-2xl ${step.bgColor} mb-4 sm:mb-6 shadow-lg mx-auto`}>
                  <step.icon className={`h-7 w-7 sm:h-9 sm:w-9 ${step.iconColor}`} />
                </div>
                <div className="text-xs font-bold tracking-widest text-gray-400 dark:text-gray-500 mb-2">STEP {step.step}</div>
                <h3 className="text-lg sm:text-xl font-bold text-gray-900 dark:text-white mb-3">
                  {step.title}
                </h3>
                <p className="text-sm sm:text-base text-gray-600 dark:text-gray-300 leading-relaxed max-w-xs mx-auto">
                  {step.description}
                </p>
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* What We Handle For You Section */}
      <section className="py-12 sm:py-16 lg:py-20 px-4 sm:px-6 lg:px-8 bg-white/50 dark:bg-gray-800/30 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto">
          <div className="text-center mb-12 sm:mb-16">
            <h2 className="text-2xl sm:text-3xl lg:text-4xl font-bold text-gray-900 dark:text-white mb-3 sm:mb-4">
              What We Handle{" "}
              <span className="bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-400 bg-clip-text text-transparent">
                For You
              </span>
            </h2>
            <p className="text-base sm:text-lg text-gray-600 dark:text-gray-300 max-w-2xl mx-auto">
              Everything that makes manual scheduling a nightmare — handled automatically.
            </p>
          </div>

          <div className="grid sm:grid-cols-2 gap-4 sm:gap-6 max-w-4xl mx-auto">
            {[
              { without: "Room double-booked by two classes", with: "Every room optimally assigned, zero overlap" },
              { without: "Faculty teaching two courses at the same time", with: "No scheduling conflicts across any faculty" },
              { without: "Days of back-and-forth manual adjustments", with: "Complete semester timetable in under 5 minutes" },
              { without: "Vague errors with no explanation", with: "System tells you exactly what it couldn't schedule and why" },
              { without: "Overloaded faculty, empty rooms", with: "Fair workload distribution, optimal room usage" },
              { without: "Starting over when one thing changes", with: "Regenerate instantly with updated input data" },
            ].map((row, index) => (
              <div
                key={index}
                className="grid grid-cols-2 gap-3 p-4 sm:p-5 rounded-xl bg-white/80 dark:bg-gray-800/50 border border-gray-200/50 dark:border-gray-700/50 shadow-sm"
              >
                <div className="flex items-start gap-2">
                  <XCircle className="h-4 w-4 text-red-400 flex-shrink-0 mt-0.5" />
                  <p className="text-xs sm:text-sm text-gray-500 dark:text-gray-400 leading-relaxed">{row.without}</p>
                </div>
                <div className="flex items-start gap-2">
                  <CheckCircle className="h-4 w-4 text-green-500 flex-shrink-0 mt-0.5" />
                  <p className="text-xs sm:text-sm text-gray-700 dark:text-gray-200 leading-relaxed font-medium">{row.with}</p>
                </div>
              </div>
            ))}

            {/* Column headers */}
            <div className="col-span-2 order-first grid grid-cols-2 gap-3 px-4 sm:px-5 pb-1">
              <p className="text-xs font-bold tracking-widest text-red-400 uppercase">Without Smart Scheduler</p>
              <p className="text-xs font-bold tracking-widest text-green-500 uppercase">With Smart Scheduler</p>
            </div>
          </div>
        </div>
      </section>

      {/* Who Is This For Section */}
      <section className="py-12 sm:py-16 lg:py-20 px-4 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto">
          <div className="text-center mb-12 sm:mb-16">
            <h2 className="text-2xl sm:text-3xl lg:text-4xl font-bold text-gray-900 dark:text-white mb-3 sm:mb-4">
              Who Is This{" "}
              <span className="bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-400 bg-clip-text text-transparent">
                For
              </span>
            </h2>
            <p className="text-base sm:text-lg text-gray-600 dark:text-gray-300 max-w-2xl mx-auto">
              One system, three different people — each getting exactly what they need.
            </p>
          </div>

          <div className="grid sm:grid-cols-3 gap-6 sm:gap-8">
            {[
              {
                icon: UserCog,
                role: "Admin / Registrar",
                hook: "Your scheduling nightmare ends here.",
                description: "Upload once. Get a complete semester timetable with zero conflicts. Approve, manage, and download — all from one dashboard.",
                iconColor: "text-blue-600 dark:text-blue-400",
                bgColor: "bg-blue-50 dark:bg-blue-900/20",
                borderColor: "border-blue-200/50 dark:border-blue-700/50",
              },
              {
                icon: BookOpen,
                role: "Faculty",
                hook: "Know your schedule before day one.",
                description: "See exactly which courses you're teaching, when and where. No surprises, no clashes, no last-minute changes nobody told you about.",
                iconColor: "text-indigo-600 dark:text-indigo-400",
                bgColor: "bg-indigo-50 dark:bg-indigo-900/20",
                borderColor: "border-indigo-200/50 dark:border-indigo-700/50",
              },
              {
                icon: GraduationCap,
                role: "Student",
                hook: "Your timetable, the moment it's ready.",
                description: "View your complete class schedule the moment it's published. Download it, save it, plan around it — done in seconds.",
                iconColor: "text-purple-600 dark:text-purple-400",
                bgColor: "bg-purple-50 dark:bg-purple-900/20",
                borderColor: "border-purple-200/50 dark:border-purple-700/50",
              },
            ].map((card, index) => (
              <Card
                key={index}
                className={`group hover:shadow-2xl transition-all duration-500 transform hover:-translate-y-2 shadow-lg bg-white/80 dark:bg-gray-800/50 backdrop-blur-sm overflow-hidden border ${card.borderColor}`}
              >
                <CardContent className="p-6 sm:p-8 text-center">
                  <div className={`inline-flex items-center justify-center w-12 h-12 sm:w-16 sm:h-16 rounded-2xl ${card.bgColor} mb-4 sm:mb-6 group-hover:scale-110 transition-all duration-500 shadow-lg`}>
                    <card.icon className={`h-6 w-6 sm:h-8 sm:w-8 ${card.iconColor}`} />
                  </div>
                  <div className="text-xs font-bold tracking-widest text-gray-400 dark:text-gray-500 uppercase mb-2">{card.role}</div>
                  <h3 className="text-lg sm:text-xl font-bold text-gray-900 dark:text-white mb-3 sm:mb-4">
                    {card.hook}
                  </h3>
                  <p className="text-sm sm:text-base text-gray-600 dark:text-gray-300 leading-relaxed">
                    {card.description}
                  </p>
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      </section>

      {/* About Section */}
      <section className="py-12 sm:py-16 lg:py-20 px-4 sm:px-6 lg:px-8 bg-white/50 dark:bg-gray-800/30 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto">
          <div className="text-center mb-12 sm:mb-16">
            <h2 className="text-2xl sm:text-3xl lg:text-4xl font-bold text-gray-900 dark:text-white mb-3 sm:mb-4">
              About{" "}
              <span className="bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-400 bg-clip-text text-transparent">
                Us
              </span>
            </h2>
            <p className="text-lg sm:text-xl text-gray-600 dark:text-gray-300 max-w-3xl mx-auto">
              Empowering you to make the most of your time.
            </p>
          </div>

          <div className="mb-12 sm:mb-16">
            <div className="relative max-w-4xl mx-auto">
              <p className="text-base sm:text-lg text-gray-600 dark:text-gray-300 leading-relaxed text-center p-6 sm:p-8 bg-white/80 dark:bg-gray-800/50 backdrop-blur-sm rounded-2xl border border-gray-200/50 dark:border-gray-700/50 shadow-lg">
                Every semester, university admins spend weeks manually juggling hundreds of courses, rooms, and faculty schedules — only to end up with conflicts anyway. We built Smart Scheduler to make that problem disappear.
              </p>
            </div>
          </div>

          <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-6 sm:gap-8">
            {[
              {
                icon: Calendar,
                title: "From Excel to Timetable in One Click",
                description: "Upload your faculty, courses, and rooms data once. The AI figures out the rest — no scheduling experience needed.",
                iconColor: "text-blue-600 dark:text-blue-400",
                bgColor: "bg-blue-50 dark:bg-blue-900/20",
              },
              {
                icon: Users,
                title: "Built for Admins, Loved by Everyone",
                description: "Admins generate, faculty view their slots, students download theirs. Everyone gets exactly what they need — nothing more, nothing less.",
                iconColor: "text-green-600 dark:text-green-400",
                bgColor: "bg-green-50 dark:bg-green-900/20",
              },
              {
                icon: Heart,
                title: "When Conflicts Happen, We Tell You Why",
                description: "If a course can't be scheduled, the system tells you exactly why — not just an error. Fix it fast, with full context.",
                iconColor: "text-purple-600 dark:text-purple-400",
                bgColor: "bg-purple-50 dark:bg-purple-900/20",
              },
            ].map((feature, index) => (
              <Card
                key={index}
                className="group hover:shadow-2xl transition-all duration-500 transform hover:-translate-y-2 border-0 shadow-lg bg-white/80 dark:bg-gray-800/50 backdrop-blur-sm overflow-hidden"
              >
                <CardContent className="p-6 sm:p-8 text-center">
                  <div className={`inline-flex items-center justify-center w-12 h-12 sm:w-16 sm:h-16 rounded-2xl ${feature.bgColor} mb-4 sm:mb-6 group-hover:scale-110 transition-all duration-500 shadow-lg`}>
                    <feature.icon className={`h-6 w-6 sm:h-8 sm:w-8 ${feature.iconColor}`} />
                  </div>
                  <h3 className="text-lg sm:text-xl font-bold text-gray-900 dark:text-white mb-3 sm:mb-4">
                    {feature.title}
                  </h3>
                  <p className="text-sm sm:text-base text-gray-600 dark:text-gray-300 leading-relaxed">
                    {feature.description}
                  </p>
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer className="bg-gray-50 dark:bg-gray-800 text-gray-900 dark:text-white py-8 sm:py-12 px-4 sm:px-6 lg:px-8 relative overflow-hidden">
        <div className="absolute inset-0 bg-gradient-to-r from-blue-600/5 to-indigo-600/5 dark:from-blue-500/5 dark:to-indigo-500/5"></div>
        <div className="max-w-7xl mx-auto text-center relative">
          <div className="flex items-center justify-center space-x-2 mb-3 sm:mb-4">
            <Calendar className="h-5 w-5 sm:h-6 sm:w-6 text-blue-600 dark:text-blue-400" />
            <span className="text-base sm:text-lg font-semibold bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-400 bg-clip-text text-transparent">
              Smart Scheduler
            </span>
          </div>
          <p className="text-sm sm:text-base text-gray-600 dark:text-gray-400">
            © 2026 Smart Scheduler. All rights reserved.
          </p>
        </div>
      </footer>
    </div>
  )
}