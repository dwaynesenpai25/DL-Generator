import { CONFIG, STATE } from "./config.js"
import { handleLarkCallback, checkSessionStatus } from "./auth.js"
import { setupMobileMenu, setupFolderSelectionButtons, setupReloadPrevention, setupNavigationHandlers } from "./ui.js"
import { setupEventHandlers } from "./event-handlers.js"
import { showSection } from "./utils.js"

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

  // Handle authentication
  const urlParams = new URLSearchParams(window.location.search)
  if (urlParams.get("code")) {
    await handleLarkCallback(urlParams.get("code"))
  } else {
    await checkSessionStatus()
  }
})

// Export global state for debugging
window.APP_STATE = STATE
window.APP_CONFIG = CONFIG
