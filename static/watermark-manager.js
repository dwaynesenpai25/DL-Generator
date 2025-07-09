// Watermark Management System
class WatermarkManager {
  constructor() {
    this.watermarkElement = null
    this.observer = null
    this.intervalId = null
    this.isActive = false
    this.userName = null
    this.userEmail = null
  }

  // Initialize watermark system
  init(userName, userEmail) {
    this.userName = userName || "Unknown User"
    this.userEmail = userEmail || ""
    this.createWatermark()
    this.setupProtection()
    this.isActive = true
  }

  // Create the watermark overlay
  createWatermark() {
    // Remove existing watermark if any
    this.removeWatermark()

    // Create watermark container
    this.watermarkElement = document.createElement("div")
    this.watermarkElement.className = "watermark-overlay"
    this.watermarkElement.id = "security-watermark"

    // Generate watermark text positions
    this.generateWatermarkText()

    // Add to document
    document.body.appendChild(this.watermarkElement)

    // Force styles to prevent tampering
    this.enforceStyles()
  }

  // Generate watermark text at multiple positions
  generateWatermarkText() {
    const container = this.watermarkElement
    const viewportWidth = window.innerWidth
    const viewportHeight = window.innerHeight

    // Calculate grid dimensions
    const textWidth = 300
    const textHeight = 200
    const cols = Math.ceil(viewportWidth / textWidth) + 1
    const rows = Math.ceil(viewportHeight / textHeight) + 1

    // Generate watermark text
    for (let row = 0; row < rows; row++) {
      for (let col = 0; col < cols; col++) {
        const textElement = document.createElement("div")
        textElement.className = "watermark-text"

        // Alternate between name and email/timestamp
        const isNameRow = row % 2 === 0
        const displayText = isNameRow ? this.userName : `${this.userEmail}`

        textElement.textContent = displayText
        textElement.style.left = `${col * textWidth - 200}px`
        textElement.style.top = `${row * textHeight - 100}px`

        container.appendChild(textElement)
      }
    }
  }

  // Enforce watermark styles to prevent tampering
  enforceStyles() {
    if (!this.watermarkElement) return

    const styles = {
      position: "fixed",
      top: "0",
      left: "0",
      width: "100vw",
      height: "100vh",
      pointerEvents: "none",
      zIndex: "9999",
      visibility: "visible",
      display: "block",
      opacity: "0.5",
    }

    Object.assign(this.watermarkElement.style, styles)
  }

  // Setup protection against tampering
  setupProtection() {
    // Monitor for DOM changes
    this.observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        // Check if watermark was removed
        if (mutation.type === "childList") {
          mutation.removedNodes.forEach((node) => {
            if (
              node === this.watermarkElement ||
              (node.nodeType === 1 && node.contains && node.contains(this.watermarkElement))
            ) {
              console.warn("Security: Watermark tampering detected - recreating")
              this.createWatermark()
            }
          })
        }

        // Check if watermark styles were modified
        if (mutation.type === "attributes" && mutation.target === this.watermarkElement) {
          console.warn("Security: Watermark style tampering detected - enforcing")
          this.enforceStyles()
        }
      })
    })

    // Start observing
    this.observer.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["style", "class"],
    })

    // Periodic integrity check
    this.intervalId = setInterval(() => {
      this.integrityCheck()
    }, 5000) // Check every 5 seconds

    // Window resize handler
    window.addEventListener("resize", () => {
      this.debounce(() => {
        this.createWatermark()
      }, 500)()
    })

    // Prevent right-click context menu on watermark
    document.addEventListener("contextmenu", (e) => {
      if (e.target.closest(".watermark-overlay")) {
        e.preventDefault()
        return false
      }
    })

    // Prevent developer tools detection (basic)
    this.setupDevToolsDetection()
  }

  // Integrity check
  integrityCheck() {
    if (!this.isActive) return

    const watermark = document.getElementById("security-watermark")

    if (!watermark) {
      console.warn("Security: Watermark missing - recreating")
      this.createWatermark()
      return
    }

    // Check if watermark is visible
    const computedStyle = window.getComputedStyle(watermark)
    if (computedStyle.display === "none" || computedStyle.visibility === "hidden" || computedStyle.opacity === "0") {
      console.warn("Security: Watermark hidden - enforcing visibility")
      this.enforceStyles()
    }

    // Check z-index
    if (Number.parseInt(computedStyle.zIndex) < 9999) {
      console.warn("Security: Watermark z-index compromised - fixing")
      watermark.style.zIndex = "9999"
    }
  }

  // Basic developer tools detection
  setupDevToolsDetection() {
    const devtools = {
      open: false,
      orientation: null,
    }

    const threshold = 160

    setInterval(() => {
      if (window.outerHeight - window.innerHeight > threshold || window.outerWidth - window.innerWidth > threshold) {
        if (!devtools.open) {
          devtools.open = true
          console.warn("Security: Developer tools detected")
          this.onDevToolsDetected()
        }
      } else {
        devtools.open = false
      }
    }, 500)
  }

  // Handle developer tools detection
  onDevToolsDetected() {
    // Refresh watermark with higher opacity
    if (this.watermarkElement) {
      const textElements = this.watermarkElement.querySelectorAll(".watermark-text")
    //   textElements.forEach((el) => {
    //     el.style.color = "rgba(255, 0, 0, 0.2)" // More visible red
    //     el.style.fontWeight = "400"
    //   })
    }
  }

  // Update watermark with new user info
  updateUser(userName, userEmail) {
    this.userName = userName || "Unknown User"
    this.userEmail = userEmail || ""
    if (this.isActive) {
      this.createWatermark()
    }
  }

  // Remove watermark (for logout)
  removeWatermark() {
    if (this.watermarkElement && this.watermarkElement.parentNode) {
      this.watermarkElement.parentNode.removeChild(this.watermarkElement)
    }
    this.watermarkElement = null
  }

  // Cleanup and destroy
  destroy() {
    this.isActive = false

    if (this.observer) {
      this.observer.disconnect()
      this.observer = null
    }

    if (this.intervalId) {
      clearInterval(this.intervalId)
      this.intervalId = null
    }

    this.removeWatermark()
  }

  // Debounce utility
  debounce(func, wait) {
    let timeout
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout)
        func(...args)
      }
      clearTimeout(timeout)
      timeout = setTimeout(later, wait)
    }
  }

  // Toggle dark mode
  setDarkMode(isDark) {
    if (this.watermarkElement) {
      if (isDark) {
        this.watermarkElement.classList.add("dark-mode")
      } else {
        this.watermarkElement.classList.remove("dark-mode")
      }
    }
  }

  // Set high contrast mode
  setHighContrast(isHighContrast) {
    if (this.watermarkElement) {
      if (isHighContrast) {
        this.watermarkElement.classList.add("high-contrast")
      } else {
        this.watermarkElement.classList.remove("high-contrast")
      }
    }
  }
}

// Create global watermark manager instance
window.watermarkManager = new WatermarkManager()

export default window.watermarkManager
