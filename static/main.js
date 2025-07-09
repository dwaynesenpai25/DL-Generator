import { CONFIG, STATE } from "./config.js"
import { handleLarkCallback, checkSessionStatus } from "./auth.js"
import { setupMobileMenu, setupFolderSelectionButtons, setupReloadPrevention, setupNavigationHandlers } from "./ui.js"
import { setupEventHandlers } from "./event-handlers.js"
import { showSection } from "./utils.js"
import watermarkManager from "./watermark-manager.js"
import watermarkSecurity from "./watermark-security.js"

// Initialize the application
document.addEventListener("DOMContentLoaded", async () => {
  // Show initial section
  showSection("dlGeneratorSection")
  document.querySelector(".card:has(#modeSelect)").classList.add("hidden")

  // Setup UI components
  setupMobileMenu()
  setupFolderSelectionButtons()
  setupReloadPrevention()
  setupNavigationHandlers()
  setupEventHandlers()

  // Initialize security systems
  watermarkSecurity.startMonitoring()

  // Handle authentication
  const urlParams = new URLSearchParams(window.location.search)
  if (urlParams.get("code")) {
    await handleLarkCallback(urlParams.get("code"))
  } else {
    await checkSessionStatus()
  }

  // Setup page visibility change handler
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) {
      console.log("Page hidden - maintaining security")
    } else {
      console.log("Page visible - verifying security")
      // Refresh watermark when page becomes visible again
      if (STATE.currentUser && watermarkManager.isActive) {
        watermarkManager.integrityCheck()
      }
    }
  })

  // Setup beforeunload handler for security
  window.addEventListener("beforeunload", (e) => {
    if (STATE.currentUser) {
      // Clean up security systems
      watermarkSecurity.stopMonitoring()
      watermarkManager.destroy()
    }
  })
})

// Export global state for debugging (in development only)
