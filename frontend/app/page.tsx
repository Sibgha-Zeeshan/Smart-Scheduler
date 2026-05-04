import Link from "next/link"
import { Calendar, Clock, Users, Heart } from "lucide-react"
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
                Welcome to Smart Scheduler, your{" "}
                <span className="font-semibold text-blue-600 dark:text-blue-400">intelligent assistant</span> for
                building a conflict-free timetable. Effortlessly organize your classes, meetings, and personal events in
                one place.
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
                <Button
                  size="lg"
                  variant="outline"
                  className="w-full sm:w-auto hover:bg-blue-50 dark:hover:bg-blue-900/20 bg-transparent border-blue-200 dark:border-blue-700"
                >
                  Learn More
                </Button>
              </div>

              {/* Stats - Hidden on mobile */}
              <div className="hidden sm:grid grid-cols-3 gap-4 lg:gap-6 pt-8 mt-8 border-t border-gray-200/50 dark:border-gray-700/50">
                {[
                  { number: "1K+", label: "Users" },
                  { number: "1K+", label: "Schedules" },
                  { number: "85-90%", label: "Efficiency" },
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
                    {
                      time: "9:00 AM",
                      subject: "Math",
                      room: "101",
                      color: "bg-gradient-to-r from-blue-500 to-blue-600 text-white",
                    },
                    {
                      time: "10:30 AM",
                      subject: "Physics",
                      room: "Lab 2",
                      color: "bg-gradient-to-r from-green-500 to-green-600 text-white",
                    },
                    {
                      time: "12:00 PM",
                      subject: "Lunch",
                      room: "Cafe",
                      color: "bg-gradient-to-r from-yellow-500 to-orange-500 text-white",
                    },
                    {
                      time: "1:00 PM",
                      subject: "Chemistry",
                      room: "Lab 1",
                      color: "bg-gradient-to-r from-purple-500 to-purple-600 text-white",
                    },
                    {
                      time: "2:30 PM",
                      subject: "English",
                      room: "205",
                      color: "bg-gradient-to-r from-pink-500 to-pink-600 text-white",
                    },
                  ].map((item, index) => (
                    <div
                      key={index}
                      className="flex items-center justify-between p-2 sm:p-3 rounded-lg border border-gray-100/50 dark:border-gray-700/50 hover:shadow-md transition-all duration-200 bg-white/50 dark:bg-gray-800/50"
                    >
                      <div className="flex items-center space-x-2 sm:space-x-3 flex-1 min-w-0">
                        <div className="text-xs sm:text-sm font-medium text-gray-600 dark:text-gray-300 w-16 sm:w-20 flex-shrink-0">
                          {item.time}
                        </div>
                        <div
                          className={`px-2 sm:px-3 py-1 rounded-lg text-xs font-medium ${item.color} truncate shadow-sm`}
                        >
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
                Smart Scheduler is built by a diverse team of educators, technologists, and dreamers who believe in the
                power of organization and simplicity. Our vision is to create a world where students and teachers can
                focus on what truly matters—learning, teaching, and personal growth—while we handle the scheduling.
              </p>
            </div>
          </div>

          <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-6 sm:gap-8">
            {[
              {
                icon: Calendar,
                title: "Easy Scheduling",
                description:
                  "Effortlessly organize your classes, meetings, and events with our intuitive scheduling tools.",
                iconColor: "text-blue-600 dark:text-blue-400",
                bgColor: "bg-blue-50 dark:bg-blue-900/20",
              },
              {
                icon: Users,
                title: "User Friendly",
                description: "Designed for everyone—students, teachers, and admins. Simple, clean, and easy to use.",
                iconColor: "text-green-600 dark:text-green-400",
                bgColor: "bg-green-50 dark:bg-green-900/20",
              },
              {
                icon: Heart,
                title: "Reliable Support",
                description:
                  "Our team is here to help you every step of the way, ensuring a smooth scheduling experience.",
                iconColor: "text-purple-600 dark:text-purple-400",
                bgColor: "bg-purple-50 dark:bg-purple-900/20",
              },
            ].map((feature, index) => (
              <Card
                key={index}
                className="group hover:shadow-2xl transition-all duration-500 transform hover:-translate-y-2 border-0 shadow-lg bg-white/80 dark:bg-gray-800/50 backdrop-blur-sm overflow-hidden"
              >
                <CardContent className="p-6 sm:p-8 text-center">
                  <div
                    className={`inline-flex items-center justify-center w-12 h-12 sm:w-16 sm:h-16 rounded-2xl ${feature.bgColor} mb-4 sm:mb-6 group-hover:scale-110 transition-all duration-500 shadow-lg`}
                  >
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
            © 2024 Smart Scheduler. All rights reserved.
          </p>
        </div>
      </footer>
    </div>
  )
}
