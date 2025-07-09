// Additional security measures for watermark protection
class WatermarkSecurity {
  constructor() {
    this.securityChecks = []
    this.isMonitoring = false
  }

  // Start comprehensive security monitoring
  startMonitoring() {
    if (this.isMonitoring) return

    this.isMonitoring = true
    this.setupConsoleProtection()
    this.setupScreenshotDetection()
    this.setupKeyboardShortcutBlocking()
    this.setupNetworkMonitoring()
  }

  // Stop security monitoring
  stopMonitoring() {
    this.isMonitoring = false
    this.securityChecks.forEach((cleanup) => cleanup())
    this.securityChecks = []
  }

  // Protect against console tampering
  setupConsoleProtection() {
    // Override console methods to detect tampering attempts
    const originalLog = console.log
    const originalWarn = console.warn
    const originalError = console.error

    console.log = (...args) => {
      const message = args.join(" ")
      if (message.includes("watermark") || message.includes("security")) {
        console.warn("Security: Console tampering detected")
      }
      return originalLog.apply(console, args)
    }

    // Cleanup function
    this.securityChecks.push(() => {
      console.log = originalLog
      console.warn = originalWarn
      console.error = originalError
    })
  }

  // Detect screenshot attempts (basic)
  setupScreenshotDetection() {
    // Monitor for common screenshot key combinations
    const screenshotKeys = [
      { key: "PrintScreen" },
      { key: "s", ctrlKey: true, shiftKey: true }, // Chrome screenshot
      { key: "F12" }, // Developer tools
      { key: "I", ctrlKey: true, shiftKey: true }, // Developer tools
      { key: "J", ctrlKey: true, shiftKey: true }, // Console
      { key: "U", ctrlKey: true }, // View source
    ]

    const keyHandler = (e) => {
      screenshotKeys.forEach((combo) => {
        if (e.key === combo.key && (!combo.ctrlKey || e.ctrlKey) && (!combo.shiftKey || e.shiftKey)) {
          console.warn("Security: Screenshot attempt detected")

          // Flash watermark to make it more visible
          this.flashWatermark()

          // Optionally prevent the action
          if (combo.key !== "PrintScreen") {
            // Can't prevent PrintScreen
            e.preventDefault()
            e.stopPropagation()
            return false
          }
        }
      })
    }

    document.addEventListener("keydown", keyHandler)

    // Cleanup function
    this.securityChecks.push(() => {
      document.removeEventListener("keydown", keyHandler)
    })
  }

  // Block certain keyboard shortcuts
  setupKeyboardShortcutBlocking() {
    const blockedCombos = [
      { key: "F12" }, // Developer tools
      { key: "I", ctrlKey: true, shiftKey: true }, // Developer tools
      { key: "J", ctrlKey: true, shiftKey: true }, // Console
      { key: "U", ctrlKey: true }, // View source
      { key: "S", ctrlKey: true }, // Save page
      { key: "A", ctrlKey: true }, // Select all (in some contexts)
    ]

    const blockHandler = (e) => {
      blockedCombos.forEach((combo) => {
        if (e.key === combo.key && (!combo.ctrlKey || e.ctrlKey) && (!combo.shiftKey || e.shiftKey)) {
          e.preventDefault()
          e.stopPropagation()

          // Show warning
          this.showSecurityWarning("Action blocked for security reasons")
          return false
        }
      })
    }

    document.addEventListener("keydown", blockHandler)

    // Cleanup function
    this.securityChecks.push(() => {
      document.removeEventListener("keydown", blockHandler)
    })
  }

  // Monitor network requests for data exfiltration
  setupNetworkMonitoring() {
    // Override fetch to monitor requests
    const originalFetch = window.fetch

    window.fetch = function (...args) {
      const url = args[0]

      // Check for suspicious external requests
      if (typeof url === "string" && !url.startsWith(window.location.origin)) {
        console.warn("Security: External request detected:", url)
      }

      return originalFetch.apply(this, args)
    }

    // Cleanup function
    this.securityChecks.push(() => {
      window.fetch = originalFetch
    })
  }

  // Flash watermark to make it more visible
  flashWatermark() {
    const watermark = document.getElementById("security-watermark")
    if (!watermark) return

    const textElements = watermark.querySelectorAll(".watermark-text")

    // Temporarily increase visibility
    textElements.forEach((el) => {
      const originalColor = el.style.color
      el.style.color = "rgba(255, 0, 0, 0.3)"
      el.style.fontWeight = "900"

      setTimeout(() => {
        el.style.color = originalColor
        el.style.fontWeight = "600"
      }, 2000)
    })
  }

  // Show security warning
  showSecurityWarning(message) {
    const warning = document.createElement("div")
    warning.className = "fixed top-4 right-4 bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded z-50"
    warning.innerHTML = `
      <div class="flex items-center gap-2">
        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L4.082 16.5c-.77.833.192 2.5 1.732 2.5z"></path>
        </svg>
        <span class="font-medium">Security Alert</span>
      </div>
      <p class="text-sm mt-1">${message}</p>
    `

    document.body.appendChild(warning)

    setTimeout(() => {
      if (warning.parentNode) {
        warning.parentNode.removeChild(warning)
      }
    }, 5000)
  }
}

// Create global security manager
window.watermarkSecurity = new WatermarkSecurity()

export default window.watermarkSecurity
