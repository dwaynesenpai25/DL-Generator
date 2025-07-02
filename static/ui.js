import { STATE } from "./config.js"
import { showError, toggleInputFields, showSection, updatePageTitle } from "./utils.js"

export function setupMobileMenu() {
  const mobileMenuToggle = document.getElementById("mobileMenuToggle")
  const mobileMenuOverlay = document.getElementById("mobileMenuOverlay")
  const sidebar = document.getElementById("sidebar")

  mobileMenuToggle.addEventListener("click", () => {
    sidebar.classList.toggle("open")
    mobileMenuOverlay.classList.toggle("hidden")
  })

  mobileMenuOverlay.addEventListener("click", () => {
    sidebar.classList.remove("open")
    mobileMenuOverlay.classList.add("hidden")
  })

  const navItems = document.querySelectorAll(".nav-item")
  navItems.forEach((item) => {
    item.addEventListener("click", () => {
      sidebar.classList.remove("open")
      mobileMenuOverlay.classList.add("hidden")
    })
  })
}

export function setupFolderSelectionButtons() {
  document.getElementById("selectAllFolders").addEventListener("click", () => {
    const checkboxes = document.querySelectorAll('#modalClients input[type="checkbox"]')
    checkboxes.forEach((checkbox) => {
      checkbox.checked = true
    })
  })

  document.getElementById("deselectAllFolders").addEventListener("click", () => {
    const checkboxes = document.querySelectorAll('#modalClients input[type="checkbox"]')
    checkboxes.forEach((checkbox) => {
      checkbox.checked = false
    })
  })
}

export function setupReloadPrevention() {
  window.addEventListener("beforeunload", (e) => {
    if (STATE.isProcessing) {
      const message =
        "Document generation is in progress. Leaving this page will cancel the process. Are you sure you want to leave?"
      e.preventDefault()
      e.returnValue = message
      return message
    }
  })

  // Prevent back/forward navigation during processing
  window.addEventListener("popstate", (e) => {
    if (STATE.isProcessing) {
      // Push the current state back to prevent navigation
      history.pushState(null, null, window.location.pathname)
      showError(
        "Cannot navigate away while document generation is in progress. Please wait for the process to complete.",
      )
    }
  })
}

export function showNoClientsModal() {
  const existingModal = document.getElementById("noClientsModal")
  if (existingModal) {
    return // Modal already exists
  }

  const modalHTML = `
    <div id="noClientsModal" class="fixed inset-0 bg-black bg-opacity-80 flex items-center justify-center z-[9999] p-4">
      <div class="glass-effect p-8 rounded-2xl shadow-large w-full max-w-md transform animate-fade-in">
        <div class="text-center mb-6">
          <div class="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
            <svg class="w-8 h-8 text-red-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                    d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L4.082 16.5c-.77.833.192 2.5 1.732 2.5z">
              </path>
            </svg>
          </div>
          <h2 class="text-2xl font-bold text-text-primary mb-2">No Access Configured</h2>
          <p class="text-text-secondary mb-4">You don't have access to any client folders yet.</p>
        </div>
        
        <div class="bg-yellow-50 border border-yellow-200 rounded-lg p-4 mb-6">
          <div class="flex items-start gap-3">
            <svg class="w-5 h-5 text-yellow-600 mt-0.5 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                    d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z">
              </path>
            </svg>
            <div>
              <p class="font-medium text-yellow-800 mb-1">What to do next:</p>
              <ul class="text-sm text-yellow-700 space-y-1">
                <li>• Contact your system administrator</li>
                <li>• Request access to the client folders you need</li>
                <li>• Wait for the administrator to configure your access</li>
              </ul>
            </div>
          </div>
        </div>
        
        <div class="text-center">
          <div class="inline-flex items-center gap-2 px-4 py-2 bg-blue-100 text-blue-800 rounded-lg text-sm">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                    d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z">
              </path>
            </svg>
            Contact Administrator for Access
          </div>
          <p class="text-xs text-text-secondary mt-2">This message will disappear once access is granted</p>
        </div>
      </div>
    </div>
  `

  document.body.insertAdjacentHTML("beforeend", modalHTML)

  // Disable all interactive elements except logout
  const interactiveElements = document.querySelectorAll(
    'button:not(#logoutButton), input, select, textarea, a:not([href="#"])',
  )
  interactiveElements.forEach((element) => {
    if (!element.closest("#noClientsModal") && element.id !== "logoutButton") {
      element.disabled = true
      element.style.opacity = "0.5"
      element.style.pointerEvents = "none"
    }
  })
}

export function hideNoClientsModal() {
  const modal = document.getElementById("noClientsModal")
  if (modal) {
    modal.remove()
    // Re-enable all interactive elements
    const interactiveElements = document.querySelectorAll("button, input, select, textarea, a")
    interactiveElements.forEach((element) => {
      element.disabled = false
      element.style.opacity = ""
      element.style.pointerEvents = ""
    })
  }
}

export function showLoginModal() {
  document.getElementById("mainContent").classList.add("hidden")
  document.getElementById("loginModal").classList.remove("hidden")
  document.getElementById("loginError").classList.add("hidden")
  hideNoClientsModal()
}

export function updateNavigationAccess() {
  const userManagementMenu = document.getElementById("userManagementMenu")
  const auditTrailMenu = document.getElementById("auditTrailMenu")

  // Check if user has admin access
  const isAdmin =
    STATE.currentUser &&
    (STATE.currentUser.access === "admin" ||
      STATE.currentUser.role === "admin" ||
      STATE.currentUser.access === "" ||
      STATE.currentUser.role === "")

  console.log("admins", STATE.currentUser)

  if (isAdmin) {
    userManagementMenu.style.display = "flex"
    auditTrailMenu.style.display = "flex"
  } else {
    userManagementMenu.style.display = "none"
    auditTrailMenu.style.display = "none"
  }
}

export function resetUI() {
  document.getElementById("outputFormatSelect").value = ""
  document.getElementById("modeSelect").value = ""
  document.getElementById("transmittalFolderSelect").value = ""
  document.querySelector(".card:has(#modeSelect)").classList.add("hidden")
  document.getElementById("transmittalFolderSection").classList.add("hidden")
  document.getElementById("printFormatInfo").classList.add("hidden")
  document.getElementById("zipFormatInfo").classList.remove("hidden")
  document.getElementById("selectionSection").classList.add("hidden")
  document.getElementById("folderSelect").innerHTML = '<option value="">Select Folder</option>'
  document.getElementById("dlTypeSelect").innerHTML = '<option value="">Select DL Type</option>'
  document.getElementById("templateSelect").innerHTML = '<option value="">Select Template</option>'
  document.getElementById("uploadSection").classList.add("hidden")
  document.getElementById("progressSection").classList.add("hidden")
  document.getElementById("errorMessage").classList.add("hidden")
  document.getElementById("placeholdersDisplay").classList.add("hidden")
  document.getElementById("dataPreview").classList.add("hidden")
  clearResultSection()
  document.getElementById("progressBar").style.width = "0%"
  document.getElementById("progressText").textContent = ""
  document.getElementById("excelUpload").value = ""
  document.getElementById("dataTable").innerHTML = ""
  document.getElementById("statusDisplay").classList.add("hidden")
  document.getElementById("templateCombinedAlert").classList.add("hidden")
  document.getElementById("excelLoadingOverlay").classList.add("hidden")

  // Clear progress details
  const progressDetails = document.getElementById("progressDetails")
  if (progressDetails) {
    progressDetails.remove()
  }

  STATE.isProcessing = false
  toggleInputFields(false)
}

export function clearResultSection() {
  const resultSection = document.getElementById("resultSection")
  resultSection.classList.add("hidden")
  const printControls = resultSection.querySelectorAll(".mt-4, .mt-6")
  printControls.forEach((control) => {
    if (
      control.querySelector("#printAreaSelect") ||
      control.querySelector("#printerSelect") ||
      control.classList.contains("mt-6")
    ) {
      control.remove()
    }
  })
  document.getElementById("downloadButton").classList.add("hidden")
  document.getElementById("cleanupButton").classList.add("hidden")
}

export function showTemplateCombinedAlert() {
  document.getElementById("templateCombinedAlert").classList.remove("hidden")
}

export function setupNavigationHandlers() {
  document.getElementById("dlGeneratorMenu").addEventListener("click", (e) => {
    e.preventDefault()
    showSection("dlGeneratorSection")
    updatePageTitle("DL Generator", "Generate and manage your documents")
  })

  document.getElementById("auditTrailMenu").addEventListener("click", (e) => {
    e.preventDefault()
    if (!STATE.currentUser || STATE.currentUser.access !== "admin") {
      showError("Access denied. Admin role required.")
      return
    }
    showSection("auditTrailSection")
    updatePageTitle("Audit Trail", "Track all document generation activities")
    // Import and call loadAuditTrail
    import("./audit-trail.js").then(({ loadAuditTrail }) => {
      loadAuditTrail()
    })
  })

  document.getElementById("userManagementMenu").addEventListener("click", (e) => {
    e.preventDefault()
    if (!STATE.currentUser || STATE.currentUser.access !== "admin") {
      showError("Access denied. Admin role required.")
      return
    }
    showSection("userManagementSection")
    updatePageTitle("User Management", "Manage user access and permissions")
    // Import and call updateUserTable
    import("./user-management.js").then(({ updateUserTable }) => {
      updateUserTable()
    })
  })
}
