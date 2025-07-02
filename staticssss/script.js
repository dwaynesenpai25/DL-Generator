let currentUser = null
const API_BASE = "http://172.20.0.86:5000/api"
// const API_BASE = "http://localhost:5000/api"
let isProcessing = false

// Initialize UI and check session
document.addEventListener("DOMContentLoaded", async () => {
  showSection("dlGeneratorSection")
  document.querySelector(".card:has(#modeSelect)").classList.add("hidden")
  setupMobileMenu()
  setupFolderSelectionButtons()
  setupReloadPrevention()

  const urlParams = new URLSearchParams(window.location.search)
  if (urlParams.get("code")) {
    await handleLarkCallback(urlParams.get("code"))
  } else {
    await checkSession()
  }
})

// Setup reload prevention during processing
function setupReloadPrevention() {
  window.addEventListener("beforeunload", (e) => {
    if (isProcessing) {
      const message =
        "Document generation is in progress. Leaving this page will cancel the process. Are you sure you want to leave?"
      e.preventDefault()
      e.returnValue = message
      return message
    }
  })

  // Prevent back/forward navigation during processing
  window.addEventListener("popstate", (e) => {
    if (isProcessing) {
      // Push the current state back to prevent navigation
      history.pushState(null, null, window.location.pathname)
      showError(
        "Cannot navigate away while document generation is in progress. Please wait for the process to complete.",
      )
    }
  })
}

// Show no-clients modal (non-removable)
function showNoClientsModal() {
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

// Hide no-clients modal and re-enable interface
function hideNoClientsModal() {
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

// Enhanced progress display with detailed stages
function updateProgressDisplay(data) {
  const progressBar = document.getElementById("progressBar")
  const progressText = document.getElementById("progressText")
  const progressSection = document.getElementById("progressSection")

  if (!progressBar || !progressText) return

  // Update progress bar
  progressBar.style.width = `${data.progress}%`

  // Enhanced progress text with stage indicators
  let stageIcon = "🔄"
  let stageColor = "text-blue-600"

  switch (data.stage) {
    case "area_processing":
      stageIcon = "🏢"
      stageColor = "text-purple-600"
      break
    case "document_generation":
      stageIcon = "📝"
      stageColor = "text-green-600"
      break
    case "transmittal_generation":
      stageIcon = "📋"
      stageColor = "text-orange-600"
      break
    case "pdf_conversion_start":
    case "conversion":
      stageIcon = "🔄"
      stageColor = "text-blue-600"
      break
    case "conversion_complete":
      stageIcon = "✅"
      stageColor = "text-green-600"
      break
    case "pdf_merging":
      stageIcon = "📑"
      stageColor = "text-indigo-600"
      break
    case "print_preparation":
      stageIcon = "📄"
      stageColor = "text-cyan-600"
      break
    case "complete":
      stageIcon = "🎉"
      stageColor = "text-green-600"
      break
    case "error":
      stageIcon = "❌"
      stageColor = "text-red-600"
      break
  }

  progressText.innerHTML = `<span class="${stageColor}">${stageIcon}</span> ${data.message}`

  // Add detailed progress information if available and not in complete stage
  if (
    (data.area_details || data.document_details || data.transmittal_details || data.conversion_details) &&
    data.stage !== "complete"
  ) {
    const detailsContainer = document.getElementById("progressDetails") || createProgressDetailsContainer()
    updateProgressDetails(detailsContainer, data)
  } else if (data.stage === "complete") {
    // Hide progress details when complete
    const detailsContainer = document.getElementById("progressDetails")
    if (detailsContainer) {
      detailsContainer.classList.add("hidden")
    }
  }
}

function createProgressDetailsContainer() {
  const progressSection = document.getElementById("progressSection")
  const detailsContainer = document.createElement("div")
  detailsContainer.id = "progressDetails"
  detailsContainer.className = "mt-4 p-4 bg-gray-50 rounded-lg border border-gray-200"

  const progressBar = document.getElementById("progressBar").parentElement
  progressBar.parentElement.insertBefore(detailsContainer, progressBar.nextSibling)

  return detailsContainer
}

function updateProgressDetails(container, data) {
  let detailsHTML = ""

  if (data.area_details) {
    detailsHTML = `
      <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">Current Area</div>
          <div class="text-lg font-bold text-purple-600">${data.area_details.current_area || "Processing..."}</div>
        </div>
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">Progress</div>
          <div class="text-lg font-bold text-blue-600">${data.area_details.area_index || 0}/${data.area_details.total_areas || 0}</div>
        </div>
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">Records in Area</div>
          <div class="text-lg font-bold text-green-600">${data.area_details.records_in_area || 0}</div>
        </div>
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">Completion</div>
          <div class="text-lg font-bold text-indigo-600">${Math.round(data.progress || 0)}%</div>
        </div>
      </div>
    `
  } else if (data.document_details) {
    detailsHTML = `
      <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">Client</div>
          <div class="text-sm font-bold text-green-600">${data.document_details.client_name || "Processing..."}</div>
        </div>
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">DL Code</div>
          <div class="text-sm font-mono font-bold text-blue-600">${data.document_details.dl_code || "N/A"}</div>
        </div>
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">Area Progress</div>
          <div class="text-lg font-bold text-purple-600">${data.document_details.record_index || 0}/${data.document_details.total_records_in_area || 0}</div>
        </div>
        <div class="bg-white p-3 rounded border">
          <div class="font-medium text-gray-600">Current Area</div>
          <div class="text-sm font-bold text-orange-600">${data.document_details.area || "Processing..."}</div>
        </div>
      </div>
    `
  } else if (data.transmittal_details) {
    detailsHTML = `
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
      <div class="bg-white p-3 rounded border">
        <div class="font-medium text-gray-600">Current Area</div>
        <div class="text-sm font-bold text-orange-600">${data.transmittal_details.area || "Processing..."}</div>
      </div>
      <div class="bg-white p-3 rounded border">
        <div class="font-medium text-gray-600">Current Page</div>
        <div class="text-lg font-bold text-purple-600">${data.transmittal_details.current_page || 0}/${data.transmittal_details.total_pages || 0}</div>
      </div>
      <div class="bg-white p-3 rounded border">
        <div class="font-medium text-gray-600">Records/Page</div>
        <div class="text-lg font-bold text-green-600">${data.transmittal_details.records_per_page || 0}</div>
      </div>
      <div class="bg-white p-3 rounded border">
        <div class="font-medium text-gray-600">Progress</div>
        <div class="text-lg font-bold text-indigo-600">${Math.round(data.progress || 0)}%</div>
      </div>
    </div>
  `
  } else if (data.conversion_details) {
    if (data.conversion_details.current_batch) {
      detailsHTML = `
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Current Batch</div>
            <div class="text-lg font-bold text-blue-600">${data.conversion_details.current_batch || 0}/${data.conversion_details.total_batches || 0}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Batch Size</div>
            <div class="text-lg font-bold text-green-600">${data.conversion_details.batch_size || 0}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Attempt</div>
            <div class="text-lg font-bold text-orange-600">${data.conversion_details.attempt || 1}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Progress</div>
            <div class="text-lg font-bold text-indigo-600">${Math.round(data.progress || 0)}%</div>
          </div>
        </div>
      `
    } else if (data.conversion_details.successful !== undefined) {
      const successRate = data.conversion_details.success_rate || 0
      const statusColor = successRate > 90 ? "text-green-600" : successRate > 70 ? "text-yellow-600" : "text-red-600"

      detailsHTML = `
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Successful</div>
            <div class="text-lg font-bold text-green-600">${data.conversion_details.successful || 0}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Failed</div>
            <div class="text-lg font-bold text-red-600">${data.conversion_details.failed || 0}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Success Rate</div>
            <div class="text-lg font-bold ${statusColor}">${successRate.toFixed(1)}%</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Batch Time</div>
            <div class="text-lg font-bold text-blue-600">${data.conversion_details.batch_time?.toFixed(1) || 0}s</div>
          </div>
        </div>
      `
    } else {
      // Basic conversion details without specific batch info
      detailsHTML = `
        <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Total Files</div>
            <div class="text-lg font-bold text-blue-600">${data.conversion_details.total_files || 0}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Total Batches</div>
            <div class="text-lg font-bold text-green-600">${data.conversion_details.total_batches || 0}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Batch Size</div>
            <div class="text-lg font-bold text-orange-600">${data.conversion_details.batch_size || 0}</div>
          </div>
          <div class="bg-white p-3 rounded border">
            <div class="font-medium text-gray-600">Progress</div>
            <div class="text-lg font-bold text-indigo-600">${Math.round(data.progress || 0)}%</div>
          </div>
        </div>
      `
    }
  }

  if (detailsHTML) {
    container.innerHTML = detailsHTML
    container.classList.remove("hidden")
  }
}

// Setup mobile menu functionality
function setupMobileMenu() {
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

function setupFolderSelectionButtons() {
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

function toggleInputFields(disabled) {
  const inputs = [
    "modeSelect",
    "folderSelect",
    "dlTypeSelect",
    "templateSelect",
    "excelUpload",
    "generateButton",
    "outputFormatSelect",
    "transmittalFolderSelect",
  ]

  inputs.forEach((id) => {
    const element = document.getElementById(id)
    if (element) {
      element.disabled = disabled
      if (disabled) {
        element.classList.add("disabled-input")
      } else {
        element.classList.remove("disabled-input")
      }
    }
  })

  const uploadLabel = document.getElementById("uploadLabel")
  if (uploadLabel) {
    if (disabled) {
      uploadLabel.classList.add("disabled-input")
    } else {
      uploadLabel.classList.remove("disabled-input")
    }
  }
}

async function checkSession() {
  try {
    const response = await fetch(`${API_BASE}/check_sessions`, {
      method: "GET",
      credentials: "include",
    })

    if (response.ok) {
      const data = await response.json()

      if (data.success) {
        currentUser = {
          username: data.username,
          role: data.role,
          access: data.access,
          clients: data.clients || [],
          userInfo: data.avatar.avatar_url || {},
        }
        document.getElementById("loginModal").classList.add("hidden")
        document.getElementById("mainContent").classList.remove("hidden")
        document.getElementById("userDisplay").textContent = `${data.username} (${data.role})`

        setUserAvatar(currentUser.userInfo)
        updateNavigationAccess()

        // Check if user has no clients and show modal if needed
        if (!currentUser.clients || currentUser.clients.length === 0) {
          showNoClientsModal()
        } else {
          hideNoClientsModal()
          fetchFolders()
        }
      } else {
        showLoginModal()
      }
    } else {
      showLoginModal()
    }
  } catch (error) {
    showLoginModal()
    console.error("Session check failed:", error)
  }
}

function setUserAvatar(userInfo) {
  const userAvatar = document.getElementById("userAvatar")
  const userAvatarFallback = document.getElementById("userAvatarFallback")

  if (userInfo) {
    userAvatar.src = userInfo
    userAvatar.classList.remove("hidden")
    userAvatarFallback.classList.add("hidden")
  } else {
    userAvatar.classList.add("hidden")
    userAvatarFallback.classList.remove("hidden")
  }
}

function updateNavigationAccess() {
  const userManagementMenu = document.getElementById("userManagementMenu")
  const auditTrailMenu = document.getElementById("auditTrailMenu")

  // Check if user has admin access (check both 'access' and 'role' properties for compatibility)
  const isAdmin =
    currentUser &&
    (currentUser.access === "admin" ||
      currentUser.role === "admin" ||
      currentUser.access === "" ||
      currentUser.role === "")
  console.log("admins", currentUser)
  if (isAdmin) {
    userManagementMenu.style.display = "flex"
    auditTrailMenu.style.display = "flex"
  } else {
    userManagementMenu.style.display = "none"
    auditTrailMenu.style.display = "none"
  }
}

let isProcessingCallback = false

async function handleLarkCallback(code) {
  if (isProcessingCallback) {
    console.log("Already processing callback, ignoring duplicate")
    return
  }
  isProcessingCallback = true

  try {
    const response = await fetch(`${API_BASE}/lark_callback?code=${code}`, {
      method: "GET",
      credentials: "include",
    })
    const data = await response.json()

    if (data.success) {
      console.log("callback", data)
      currentUser = { username: data.username, role: data.role }
      document.getElementById("loginModal").classList.add("hidden")
      document.getElementById("mainContent").classList.remove("hidden")
      document.getElementById("userDisplay").textContent = `${data.username} (${data.role})`
      window.history.replaceState({}, document.title, "/")
      await checkSession()
      updateNavigationAccess()
    } else {
      document.getElementById("loginError").classList.remove("hidden")
      document.getElementById("loginError").textContent = data.detail || "Authentication failed."
    }
  } catch (error) {
    console.error("Error in handleLarkCallback:", error)
    document.getElementById("loginError").classList.remove("hidden")
    document.getElementById("loginError").textContent = "Authentication failed. Please try again."
  } finally {
    isProcessingCallback = false
  }
}

document.getElementById("loginButton").addEventListener("click", () => {
  window.location.href = `${API_BASE}/login`
})

document.getElementById("logoutButton").addEventListener("click", async () => {
  try {
    const response = await fetch(`${API_BASE}/logout`, {
      method: "GET",
      credentials: "include",
    })
    const data = await response.json()
    if (data.success) {
      currentUser = null
      showLoginModal()
      resetUI()
      hideNoClientsModal() // Hide the no-clients modal on logout
      window.history.replaceState({}, document.title, "/")
    } else {
      showError(data.detail || "Logout failed.")
    }
  } catch (error) {
    showError("Logout failed. Please check if the backend server is running.")
  }
})

function showLoginModal() {
  document.getElementById("mainContent").classList.add("hidden")
  document.getElementById("loginModal").classList.remove("hidden")
  document.getElementById("loginError").classList.add("hidden")
  hideNoClientsModal() // Hide no-clients modal when showing login
}

function redirectToLogin() {
  currentUser = null
  showLoginModal()
  resetUI()
  showError("Session expired. Please log in again.")
}

function updatePageTitle(title, subtitle) {
  document.getElementById("pageTitle").textContent = title
  document.getElementById("pageSubtitle").textContent = subtitle
}

// Navigation handlers
document.getElementById("dlGeneratorMenu").addEventListener("click", (e) => {
  e.preventDefault()
  showSection("dlGeneratorSection")
  updatePageTitle("DL Generator", "Generate and manage your documents")
})

document.getElementById("auditTrailMenu").addEventListener("click", (e) => {
  e.preventDefault()
  if (!currentUser || currentUser.access !== "admin") {
    showError("Access denied. Admin role required.")
    return
  }
  showSection("auditTrailSection")
  updatePageTitle("Audit Trail", "Track all document generation activities")
  loadAuditTrail()
})

document.getElementById("userManagementMenu").addEventListener("click", (e) => {
  e.preventDefault()
  if (!currentUser || currentUser.access !== "admin") {
    showError("Access denied. Admin role required.")
    return
  }
  showSection("userManagementSection")
  updatePageTitle("User Management", "Manage user access and permissions")
  updateUserTable()
})

function showSection(sectionId) {
  document.getElementById("dlGeneratorSection").classList.add("hidden")
  document.getElementById("userManagementSection").classList.add("hidden")
  document.getElementById("auditTrailSection").classList.add("hidden")
  document.getElementById(sectionId).classList.remove("hidden")

  const navItems = document.querySelectorAll(".nav-item")
  navItems.forEach((item) => {
    item.classList.remove("active")
  })

  if (sectionId === "dlGeneratorSection") {
    document.getElementById("dlGeneratorMenu").classList.add("active")
  } else if (sectionId === "userManagementSection") {
    document.getElementById("userManagementMenu").classList.add("active")
  } else if (sectionId === "auditTrailSection") {
    document.getElementById("auditTrailMenu").classList.add("active")
  }
}

function resetUI() {
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

  isProcessing = false
  toggleInputFields(false)
}

function clearResultSection() {
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

function showError(message, isCritical = false) {
  const errorDiv = document.getElementById("errorMessage")
  const generalErrorContainer = document.getElementById("generalErrorContainer")

  let displayMessage = message || "An unexpected error occurred. Please try again."
  if (typeof message === "object" && message !== null) {
    if (message.detail) displayMessage = message.detail
    else if (message.message) displayMessage = message.message
    else displayMessage = JSON.stringify(message)
  }

  if (errorDiv && errorDiv.offsetParent !== null) {
    const errorTextSpan = errorDiv.querySelector("#errorText")
    if (errorTextSpan) {
      errorTextSpan.textContent = displayMessage
      errorDiv.classList.remove("hidden")
    } else {
      errorDiv.innerHTML = `<div class="bg-red-50 border border-red-200 rounded-xl p-4 animate-fade-in">...${displayMessage}</div>`
      errorDiv.classList.remove("hidden")
    }
  } else if (generalErrorContainer) {
    generalErrorContainer.innerHTML = `
      <div class="fixed top-4 right-4 bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded-md shadow-lg z-50 animate-fade-in" role="alert">
        <div class="flex">
          <div class="py-1"><svg class="fill-current h-6 w-6 text-red-500 mr-4" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20"><path d="M2.93 17.07A10 10 0 1 17.07 2.93 10 10 0 0 1 2.93 17.07zM11.414 10l2.829-2.828-1.415-1.415L10 8.586 7.172 5.757 5.757 7.172 8.586 10l-2.829 2.828 1.415 1.415L10 11.414l2.828 2.829 1.415-1.415L11.414 10z"/></svg></div>
          <div>
            <p class="font-bold">Error</p>
            <p class="text-sm">${displayMessage}</p>
          </div>
        </div>
      </div>
    `
    generalErrorContainer.classList.remove("hidden")
    setTimeout(() => {
      generalErrorContainer.classList.add("hidden")
      generalErrorContainer.innerHTML = ""
    }, 5000)
  } else {
    console.error("Error display UI element not found. Raw error:", displayMessage)
  }

  if (isCritical) {
    console.warn("A critical error occurred:", displayMessage)
    toggleInputFields(true)
  }
}

function showTemplateCombinedAlert() {
  document.getElementById("templateCombinedAlert").classList.remove("hidden")
}

// Fetch folders (now user-specific)
async function fetchFolders() {
  try {
    const response = await fetch(`${API_BASE}/folders`, { credentials: "include" })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin() // This already handles UI reset
        return
      }
      console.log("1", response)
      const errorData = await response
        .json()
        .catch(() => ({ detail: "Failed to fetch folders and parse error response." }))
      throw new Error(errorData.detail || `HTTP error ${response.status}`)
    }
    const folders = await response.json()
    const folderSelect = document.getElementById("folderSelect")
    folderSelect.innerHTML = '<option value="">Select Folder</option>'
    folders.forEach((folder) => {
      const option = document.createElement("option")
      option.value = folder
      option.textContent = folder
      folderSelect.appendChild(option)
    })

    if (currentUser && currentUser.access === "admin1") {
      await loadTemplateFoldersForModal()
    }
  } catch (error) {
    console.error("fetchFolders error:", error)
    showError(`Failed to fetch folders: ${error.message}. Please check FTP configuration or server status.`)
  }
}

// Enhanced generate button with improved progress handling and reload prevention
document.getElementById("generateButton").addEventListener("click", async () => {
  const file = document.getElementById("excelUpload").files[0]
  if (!file) {
    showError("Please upload an Excel file")
    return
  }

  // Set processing state and prevent reloads
  isProcessing = true
  toggleInputFields(true)

  // Push current state to prevent back navigation
  history.pushState(null, null, window.location.pathname)

  document.getElementById("progressSection").classList.remove("hidden")
  document.getElementById("errorMessage").classList.add("hidden")
  const progressBar = document.getElementById("progressBar")
  const progressText = document.getElementById("progressText")
  const resultSection = document.getElementById("resultSection")
  const downloadButton = document.getElementById("downloadButton")
  const cleanupButton = document.getElementById("cleanupButton")

  progressBar.style.width = "0%"
  progressText.innerHTML = '<span class="text-blue-600">🚀</span> Starting processing...'

  clearResultSection()

  const formData = new FormData()
  formData.append("file", file)

  let timeoutId
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = setTimeout(() => {
      reject(new Error("Processing timed out after 120 seconds"))
    }, 120000)
  })

  try {
    const response = await Promise.race([
      fetch(`${API_BASE}/generate_pdfs`, {
        method: "POST",
        body: formData,
        credentials: "include",
      }),
      timeoutPromise,
    ])

    clearTimeout(timeoutId)

    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Server error")
    }

    const reader = response.body.getReader()
    const decoder = new TextDecoder()

    while (true) {
      const { done, value } = await reader.read()
      if (done) break

      try {
        const chunk = decoder.decode(value, { stream: true })
        const jsonObjects = chunk.split("\n").filter((line) => line.trim())

        for (const jsonStr of jsonObjects) {
          try {
            const data = JSON.parse(jsonStr)

            if (data.error) {
              showError(data.error)
              return
            }

            // Use enhanced progress display
            updateProgressDisplay(data)

            if (data.download_ready) {
              resultSection.classList.remove("hidden")
              downloadButton.classList.remove("hidden")
              cleanupButton.classList.remove("hidden")
              downloadButton.onclick = () => {
                window.location.href = `${API_BASE}/download_zip`
              }
            }

            // Handle print-ready response with enhanced UI
            if (data.print_ready) {
              resultSection.classList.remove("hidden")
              cleanupButton.classList.remove("hidden")

              // Load available printers
              const printers = await loadAvailablePrinters()

              // Create enhanced print controls
              const printContainer = createPrintControls(data.areas, printers)
              resultSection.appendChild(printContainer)
            }
          } catch (jsonError) {
            console.error("Error parsing JSON chunk:", jsonError, jsonStr)
          }
        }
      } catch (chunkError) {
        console.error("Error processing chunk:", chunkError)
      }
    }
  } catch (error) {
    clearTimeout(timeoutId)
    showError(`Failed to generate PDFs: ${error.message}`)
    progressText.innerHTML = '<span class="text-red-600">❌</span> Processing failed.'
    progressBar.style.width = "0%"
  } finally {
    // Reset processing state
    isProcessing = false
    toggleInputFields(false)
  }
})

// Rest of the existing functions remain the same...
// [Include all the remaining functions from the original script.js file]

// Load template folders for user modal
async function loadTemplateFoldersForModal() {
  try {
    // Use the new endpoint that returns all folders for admin users
    const response = await fetch(`${API_BASE}/all_folders`, {
      credentials: "include",
    })
    if (response.ok) {
      const folders = await response.json()
      const modalClients = document.getElementById("modalClients")
      modalClients.innerHTML = ""

      folders.forEach((folder) => {
        const checkboxDiv = document.createElement("div")
        checkboxDiv.className = "flex items-center"
        checkboxDiv.innerHTML = `
          <input type="checkbox" id="client_${folder}" value="${folder}" 
                 class="mr-3 w-4 h-4 text-primary bg-surface border-border-medium rounded focus:ring-primary focus:ring-2">
          <label for="client_${folder}" class="text-sm text-text-primary cursor-pointer">${folder}</label>
        `
        modalClients.appendChild(checkboxDiv)
      })
    } else {
      console.error("Failed to load template folders:", response.status)
    }
  } catch (error) {
    console.error("Failed to load template folders for modal:", error)
  }
}

// Fetch DL types
async function fetchDLTypes(folder) {
  try {
    const response = await fetch(`${API_BASE}/dl_types`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ folder }),
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      if (response.status === 403) {
        showError("Access denied to this template folder.")
        return
      }
      const errorData = await response
        .json()
        .catch(() => ({ detail: "Failed to fetch DL types and parse error response." }))
      throw new Error(errorData.detail || `HTTP error ${response.status}`)
    }
    const dlTypes = await response.json()
    const dlTypeSelect = document.getElementById("dlTypeSelect")
    dlTypeSelect.innerHTML = '<option value="">Select DL Type</option>'
    if (dlTypes && dlTypes.length > 0) {
      dlTypes.forEach((type) => {
        const option = document.createElement("option")
        option.value = type
        option.textContent = type
        dlTypeSelect.appendChild(option)
      })
    } else {
      showError(`No DL types found for folder "${folder}". Check Google Sheets configuration.`)
    }
  } catch (error) {
    console.error("fetchDLTypes error:", error)
    showError(`Failed to fetch DL types for "${folder}": ${error.message}.`)
  }
}

// Fetch templates
async function fetchTemplates(folder) {
  try {
    const response = await fetch(`${API_BASE}/templates`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ folder }),
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      if (response.status === 403) {
        showError("Access denied to this template folder.")
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || `HTTP error: ${response.status}`)
    }
    const data = await response.json()
    const templates = data.templates || []
    if (!Array.isArray(templates)) {
      throw new Error("Invalid templates format received from server.")
    }
    const templateSelect = document.getElementById("templateSelect")
    templateSelect.innerHTML = '<option value="">Select Template</option>'
    templates.forEach((template) => {
      const option = document.createElement("option")
      option.value = template
      option.textContent = template
      templateSelect.appendChild(option)
    })
  } catch (error) {
    showError(`Failed to fetch templates: ${error.message}. Please check FTP connection or server status.`)
  }
}

// Fetch placeholders
async function fetchPlaceholders(folder, dl_type, template) {
  document.getElementById("placeholdersLoadingOverlay").classList.remove("hidden")
  try {
    const response = await fetch(`${API_BASE}/placeholders`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ folder, dl_type, template }),
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      if (response.status === 403) {
        showError("Access denied to this template folder.")
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to fetch placeholders")
    }
    const data = await response.json()

    const placeholdersList = document.getElementById("placeholdersList")
    if (data.message) {
      placeholdersList.innerHTML = ""
      const placeholders = data.placeholders || []
      // Filter out placeholders starting with "IMAGE_" and clean «»
      const filteredPlaceholders = placeholders
        .filter((placeholder) => !placeholder.startsWith("«IMAGE_"))
        .map((placeholder) => placeholder.replace(/«|»/g, ""))

      if (filteredPlaceholders.length > 0) {
        filteredPlaceholders.forEach((placeholder) => {
          const li = document.createElement("li")
          li.className = "flex items-center gap-2 text-sm text-text-secondary"
          li.innerHTML = `
            <div class="w-2 h-2 bg-accent rounded-full"></div>
            <code class="bg-background px-2 py-1 rounded text-xs font-mono">${placeholder}</code>
          `
          placeholdersList.appendChild(li)
        })
      } else {
        placeholdersList.innerHTML = '<li class="text-sm text-text-secondary">No valid placeholders found.</li>'
      }

      document.getElementById("placeholdersDisplay").classList.remove("hidden")
      document.getElementById("uploadSection").classList.remove("hidden")

      // Check if template is already combined and show alert
      if (data.template_combined === true) {
        showTemplateCombinedAlert()
      }
    } else {
      document.getElementById("templatecontentStatusText").textContent = data.detail
      document.getElementById("templatecontentStatus").classList.remove("hidden")
    }
  } catch (error) {
    console.error("Error processing placeholders:", error)
    showError("Failed to fetch placeholders. Please check template configuration.")
  } finally {
    // Hide loading overlay
    document.getElementById("placeholdersLoadingOverlay").classList.add("hidden")
  }
}

async function fetchTransmittalPlaceholders() {
  document.getElementById("placeholdersLoadingOverlay").classList.remove("hidden")
  try {
    const response = await fetch(`${API_BASE}/transmittal_placeholders`, {
      method: "GET",
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to fetch transmittal placeholders")
    }
    const data = await response.json()

    const placeholdersList = document.getElementById("placeholdersList")
    if (data.message) {
      placeholdersList.innerHTML = ""
      const placeholders = data.placeholders || []
      // Filter out placeholders starting with "IMAGE_" and clean «»
      const filteredPlaceholders = placeholders
        .filter((placeholder) => !placeholder.startsWith("«IMAGE_"))
        .map((placeholder) => placeholder.replace(/«|»/g, ""))

      if (filteredPlaceholders.length > 0) {
        filteredPlaceholders.forEach((placeholder) => {
          const li = document.createElement("li")
          li.className = "flex items-center gap-2 text-sm text-text-secondary"
          li.innerHTML = `
            <div class="w-2 h-2 bg-accent rounded-full"></div>
            <code class="bg-background px-2 py-1 rounded text-xs font-mono">${placeholder}</code>
          `
          placeholdersList.appendChild(li)
        })
      } else {
        placeholdersList.innerHTML = '<li class="text-sm text-text-secondary">No valid placeholders found.</li>'
      }

      document.getElementById("placeholdersDisplay").classList.remove("hidden")
      document.getElementById("uploadSection").classList.remove("hidden")
    } else {
      document.getElementById("templateStatusText").textContent = data.detail
      document.getElementById("statusDisplay").classList.remove("hidden")
    }
  } catch (error) {
    console.error("Error processing transmittal placeholders:", error)
    showError("Failed to fetch transmittal placeholders. Please check template configuration.")
  } finally {
    // Hide loading overlay
    document.getElementById("placeholdersLoadingOverlay").classList.add("hidden")
  }
}

// Global variables for pagination
let currentAuditPage = 1
let currentAuditDetailsPage = 1

// Load audit trail with pagination
async function loadAuditTrail(page = 1) {
  try {
    currentAuditPage = page
    const response = await fetch(`${API_BASE}/audit_trail?page=${page}&limit=10`, {
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      throw new Error("Failed to fetch audit trail")
    }
    const data = await response.json()
    const tbody = document.getElementById("auditTableBody")
    tbody.innerHTML = ""

    if (data.entries.length === 0) {
      tbody.innerHTML =
        '<tr><td colspan="5" class="px-6 py-12 text-center text-text-secondary"><div class="flex flex-col items-center gap-3"><svg class="w-12 h-12 text-neutral-light" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v10a2 2 0 002 2h8a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path></svg><p class="text-lg font-medium">No audit entries found</p><p class="text-sm">Start generating documents to see audit trail</p></div></td></tr>'
      return
    }

    data.entries.forEach((entry, index) => {
      const row = document.createElement("tr")
      row.className = "clickable-row transition-all duration-200 ease-in-out"
      row.onclick = () => showAuditDetails(entry.id)

      // Add alternating row colors for better readability
      if (index % 2 === 0) {
        row.classList.add("bg-gray-50")
      }

      const processedDate = new Date(entry.processed_at)
      const formattedDate = processedDate.toLocaleDateString("en-US", {
        year: "numeric",
        month: "short",
        day: "numeric",
      })
      const formattedTime = processedDate.toLocaleTimeString("en-US", {
        hour: "2-digit",
        minute: "2-digit",
      })

      row.innerHTML = `
                <td class="px-6 py-4">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 bg-primary bg-opacity-10 rounded-lg flex items-center justify-center">
                            <svg class="w-5 h-5 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"></path>
                            </svg>
                        </div>
                        <div>
                            <p class="font-semibold text-text-primary">${entry.client}</p>
                            <p class="text-xs text-text-secondary">Click to view details</p>
                        </div>
                    </div>
                </td>
                <td class="px-6 py-4">
                    <div class="flex items-center gap-2">
                        <div class="w-8 h-8 bg-secondary bg-opacity-10 rounded-full flex items-center justify-center">
                            <svg class="w-4 h-4 text-secondary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path>
                            </svg>
                        </div>
                        <span class="text-text-secondary font-medium">${entry.processed_by}</span>
                    </div>
                </td>
                <td class="px-6 py-4">
                    <div class="text-text-secondary">
                        <p class="font-medium">${formattedDate}</p>
                        <p class="text-xs text-neutral-light">${formattedTime}</p>
                    </div>
                </td>
                <td class="px-6 py-4">
                    <div class="flex items-center gap-2">
                        <div class="w-8 h-8 bg-accent bg-opacity-10 rounded-full flex items-center justify-center">
                            <svg class="w-4 h-4 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path>
                            </svg>
                        </div>
                        <span class="text-text-secondary font-bold text-lg">${entry.total_accounts}</span>
                    </div>
                </td>
                <td class="px-6 py-4">
                    <span class="inline-flex items-center px-3 py-1 text-xs font-semibold rounded-full ${getModeColor(entry.mode)}">
                        ${entry.mode}
                    </span>
                </td>
            `
      tbody.appendChild(row)
    })

    // Render pagination
    renderAuditPagination(data.pagination)
  } catch (error) {
    showError("Failed to load audit trail.")
  }
}

// Render audit trail pagination
function renderAuditPagination(pagination) {
  const paginationContainer = document.getElementById("auditPagination")
  if (!paginationContainer) return

  let paginationHTML = `
    <div class="flex items-center justify-between">
      <div class="text-sm text-gray-700">
        Showing page ${pagination.current_page} of ${pagination.total_pages} 
        (${pagination.total_count} total entries)
      </div>
      <div class="flex items-center gap-2">
  `

  // Previous button
  if (pagination.has_prev) {
    paginationHTML += `
      <button onclick="loadAuditTrail(${pagination.current_page - 1})" 
              class="px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
        Previous
      </button>
    `
  } else {
    paginationHTML += `
      <button disabled class="px-3 py-2 text-sm font-medium text-gray-300 bg-gray-100 border border-gray-200 rounded-md cursor-not-allowed">
        Previous
      </button>
    `
  }

  // Page numbers
  const startPage = Math.max(1, pagination.current_page - 2)
  const endPage = Math.min(pagination.total_pages, pagination.current_page + 2)

  for (let i = startPage; i <= endPage; i++) {
    if (i === pagination.current_page) {
      paginationHTML += `
        <button class="px-3 py-2 text-sm font-medium text-white bg-primary border border-primary rounded-md">
          ${i}
        </button>
      `
    } else {
      paginationHTML += `
        <button onclick="loadAuditTrail(${i})" 
                class="px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
          ${i}
        </button>
      `
    }
  }

  // Next button
  if (pagination.has_next) {
    paginationHTML += `
      <button onclick="loadAuditTrail(${pagination.current_page + 1})" 
              class="px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
        Next
      </button>
    `
  } else {
    paginationHTML += `
      <button disabled class="px-3 py-2 text-sm font-medium text-gray-300 bg-gray-100 border border-gray-200 rounded-md cursor-not-allowed">
        Next
      </button>
    `
  }

  paginationHTML += `
      </div>
    </div>
  `

  paginationContainer.innerHTML = paginationHTML
}

// Helper function to get mode-specific colors
function getModeColor(mode) {
  switch (mode) {
    case "DL Only":
      return "bg-blue-100 text-blue-800"
    case "DL w/ Transmittal":
      return "bg-purple-100 text-purple-800"
    case "Transmittal Only":
      return "bg-green-100 text-green-800"
    default:
      return "bg-gray-100 text-gray-800"
  }
}

// Enhanced audit details modal with pagination
async function showAuditDetails(auditId, page = 1) {
  try {
    currentAuditDetailsPage = page

    // Close any existing modal first
    const existingModal = document.getElementById("auditDetailsModal")
    if (existingModal) {
      document.body.removeChild(existingModal)
    }

    // Show loading indicator
    const loadingModal = document.createElement("div")
    loadingModal.id = "loadingModal"
    loadingModal.className = "fixed inset-0 bg-black bg-opacity-70 flex items-center justify-center z-50 p-4"
    loadingModal.innerHTML = `
      <div class="glass-effect p-8 rounded-2xl shadow-large flex items-center gap-4">
        <div class="loading-spinner"></div>
        <div>
          <p class="text-text-primary font-semibold text-lg">Loading Account Details</p>
          <p class="text-text-secondary text-sm">Please wait while we fetch the data...</p>
        </div>
      </div>
    `
    document.body.appendChild(loadingModal)

    // Fetch processed accounts for this audit entry
    const response = await fetch(`${API_BASE}/audit_details/${auditId}?page=${page}&limit=50`, {
      credentials: "include",
    })

    // Remove loading indicator
    if (document.getElementById("loadingModal")) {
      document.body.removeChild(loadingModal)
    }

    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      throw new Error("Failed to fetch audit details")
    }

    const details = await response.json()

    // Create enhanced modal content
    const modalContent = document.createElement("div")
    modalContent.className =
      "glass-effect p-8 rounded-2xl shadow-large w-full max-w-6xl transform animate-fade-in max-h-[90vh] overflow-hidden flex flex-col"
    modalContent.innerHTML = `
      <div class="flex items-center justify-between mb-6 flex-shrink-0">
        <div class="flex items-center gap-4">
          <div class="w-12 h-12 bg-primary bg-opacity-10 rounded-xl flex items-center justify-center">
            <svg class="w-6 h-6 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v10a2 2 0 002 2h8a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path>
            </svg>
          </div>
          <div>
            <h2 class="text-2xl font-bold text-text-primary">Processed Accounts</h2>
            <p class="text-text-secondary">Audit ID: ${details.audit_id}</p>
          </div>
        </div>
        <button id="closeAuditModal" class="p-3 rounded-xl hover:bg-surface-hover transition-colors">
          <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
          </svg>
        </button>
      </div>
      
      <div class="grid grid-cols-2 lg:grid-cols-4 gap-6 mb-6 flex-shrink-0">
        <div class="bg-blue-50 p-4 rounded-xl border border-blue-200">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-blue-100 rounded-lg flex items-center justify-center">
              <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"></path>
              </svg>
            </div>
            <div>
              <p class="text-sm text-blue-600 font-medium">Client</p>
              <p class="font-bold text-blue-900">${details.client}</p>
            </div>
          </div>
        </div>
        <div class="bg-green-50 p-4 rounded-xl border border-green-200">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-green-100 rounded-lg flex items-center justify-center">
              <svg class="w-5 h-5 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path>
              </svg>
            </div>
            <div>
              <p class="text-sm text-green-600 font-medium">Processed By</p>
              <p class="font-bold text-green-900">${details.processed_by}</p>
            </div>
          </div>
        </div>
        <div class="bg-purple-50 p-4 rounded-xl border border-purple-200">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-purple-100 rounded-lg flex items-center justify-center">
              <svg class="w-5 h-5 text-purple-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path>
              </svg>
            </div>
            <div>
              <p class="text-sm text-purple-600 font-medium">Date</p>
              <p class="font-bold text-purple-900">${new Date(details.processed_at).toLocaleDateString()}</p>
            </div>
          </div>
        </div>
        <div class="bg-orange-50 p-4 rounded-xl border border-orange-200">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-orange-100 rounded-lg flex items-center justify-center">
              <svg class="w-5 h-5 text-orange-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
              </svg>
            </div>
            <div>
              <p class="text-sm text-orange-600 font-medium">Mode</p>
              <p class="font-bold text-orange-900">${details.mode}</p>
            </div>
          </div>
        </div>
      </div>
      
      <div class="flex-1 overflow-hidden">
        <div class="flex items-center justify-between mb-4">
          <h3 class="text-lg font-semibold text-text-primary">Account Details</h3>
          <div class="text-sm text-text-secondary">
            Showing ${details.accounts.length} of ${details.pagination.total_count} accounts
          </div>
        </div>
        
        <div class="overflow-auto rounded-lg border border-border-light" style="height: 400px;">
          <table class="min-w-full overflow-auto">
            <thead class="sticky top-0 bg-white z-10">
              <tr class="border-b border-border-light bg-gradient-to-r from-gray-50 to-gray-100">
                <th class="text-left py-4 px-6 font-semibold text-text-primary">
                  <div class="flex items-center gap-2">
                    <svg class="w-4 h-4 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 20l4-16m2 16l4-16M6 9h14M4 15h14"></path>
                    </svg>
                    DL Code
                  </div>
                </th>
                <th class="text-left py-4 px-6 font-semibold text-text-primary">
                  <div class="flex items-center gap-2">
                    <svg class="w-4 h-4 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path>
                    </svg>
                    Name
                  </div>
                </th>
                <th class="text-left py-4 px-6 font-semibold text-text-primary">
                  <div class="flex items-center gap-2">
                    <svg class="w-4 h-4 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17.657 16.657L13.414 20.9a1.998 1.998 0 01-2.827 0l-4.244-4.243a8 8 0 1111.314 0z"></path>
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 11a3 3 0 11-6 0 3 3 0 016 0z"></path>
                    </svg>
                    Address
                  </div>
                </th>
                <th class="text-left py-4 px-6 font-semibold text-text-primary">
                  <div class="flex items-center gap-2">
                    <svg class="w-4 h-4 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 20l-5.447-2.724A1 1 0 013 16.382V5.618a1 1 0 011.447-.894L9 7m0 13l6-3m-6 3V7m6 10l4.553 2.276A1 1 0 0021 18.382V7.618a1 1 0 00-.553-.894L15 4m0 13V4m0 0L9 7"></path>
                    </svg>
                    Area
                  </div>
                </th>
              </tr>
            </thead>
            <tbody id="accountDetailsList" class="divide-y divide-border-light">
              ${
                details.accounts.length > 0
                  ? details.accounts
                      .map(
                        (account, index) => `
                  <tr class="hover:bg-surface-hover transition-colors ${index % 2 === 0 ? "bg-gray-50" : "bg-white"}">
                    <td class="px-6 py-4">
                      <div class="flex items-center gap-3">
                        <div class="w-8 h-8 bg-primary bg-opacity-10 rounded-lg flex items-center justify-center">
                          <svg class="w-4 h-4 text-primary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 20l4-16m2 16l4-16M6 9h14M4 15h14"></path>
                          </svg>
                        </div>
                        <span class="font-mono text-sm font-medium text-text-primary">${account.dl_code || "N/A"}</span>
                      </div>
                    </td>
                    <td class="px-6 py-4">
                      <span class="font-medium text-text-primary">${account.name || "N/A"}</span>
                    </td>
                    <td class="px-6 py-4">
                      <span class="text-text-secondary text-sm">${account.address || "N/A"}</span>
                    </td>
                    <td class="px-6 py-4">
                      <span class="inline-flex items-center px-2 py-1 text-xs font-medium bg-accent bg-opacity-10 text-accent rounded-full">
                        ${account.area || "N/A"}
                      </span>
                    </td>
                  </tr>
                `,
                      )
                      .join("")
                  : '<tr><td colspan="4" class="px-6 py-12 text-center text-text-secondary"><div class="flex flex-col items-center gap-3"><svg class="w-12 h-12 text-neutral-light" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v10a2 2 0 002 2h8a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path></svg><p class="text-lg font-medium">No account details available</p><p class="text-sm">The processed accounts data could not be found</p></div></td></tr>'
              }
            </tbody>
          </table>
        </div>
        
        <!-- Modal Pagination -->
        <div id="modalPagination" class="mt-4">
          ${renderModalPagination(details.pagination, auditId)}
        </div>
      </div>
    `

    // Create modal container
    const modalContainer = document.createElement("div")
    modalContainer.id = "auditDetailsModal"
    modalContainer.className = "fixed inset-0 bg-black bg-opacity-70 flex items-center justify-center z-50 p-4"
    modalContainer.appendChild(modalContent)

    // Add to document
    document.body.appendChild(modalContainer)

    // Add close event
    document.getElementById("closeAuditModal").addEventListener("click", () => {
      const modal = document.getElementById("auditDetailsModal")
      if (modal) {
        document.body.removeChild(modal)
      }
    })

    // Close on outside click
    modalContainer.addEventListener("click", (e) => {
      if (e.target === modalContainer) {
        const modal = document.getElementById("auditDetailsModal")
        if (modal) {
          document.body.removeChild(modal)
        }
      }
    })
  } catch (error) {
    console.error("Failed to load audit details:", error)
    showError("Failed to load audit details.")

    // Clean up any loading modal
    const loadingModal = document.getElementById("loadingModal")
    if (loadingModal) {
      document.body.removeChild(loadingModal)
    }
  }
}

// Render modal pagination
function renderModalPagination(pagination, auditId) {
  let paginationHTML = `
    <div class="flex items-center justify-between">
      <div class="text-sm text-gray-700">
        Page ${pagination.current_page} of ${pagination.total_pages} 
        (${pagination.total_count} total accounts)
      </div>
      <div class="flex items-center gap-2">
  `

  // Previous button
  if (pagination.has_prev) {
    paginationHTML += `
      <button onclick="showAuditDetails(${auditId}, ${pagination.current_page - 1})" 
              class="px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
        Previous
      </button>
    `
  } else {
    paginationHTML += `
      <button disabled class="px-3 py-2 text-sm font-medium text-gray-300 bg-gray-100 border border-gray-200 rounded-md cursor-not-allowed">
        Previous
      </button>
    `
  }

  // Page numbers
  const startPage = Math.max(1, pagination.current_page - 2)
  const endPage = Math.min(pagination.total_pages, pagination.current_page + 2)

  for (let i = startPage; i <= endPage; i++) {
    if (i === pagination.current_page) {
      paginationHTML += `
        <button class="px-3 py-2 text-sm font-medium text-white bg-primary border border-primary rounded-md">
          ${i}
        </button>
      `
    } else {
      paginationHTML += `
        <button onclick="showAuditDetails(${auditId}, ${i})" 
                class="px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
          ${i}
        </button>
      `
    }
  }

  // Next button
  if (pagination.has_next) {
    paginationHTML += `
      <button onclick="showAuditDetails(${auditId}, ${pagination.current_page + 1})" 
              class="px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-md hover:bg-gray-50">
        Next
      </button>
    `
  } else {
    paginationHTML += `
      <button disabled class="px-3 py-2 text-sm font-medium text-gray-300 bg-gray-100 border border-gray-200 rounded-md cursor-not-allowed">
        Next
      </button>
    `
  }

  paginationHTML += `
      </div>
    </div>
  `

  return paginationHTML
}

async function deleteUser(email) {
  if (!confirm(`Are you sure you want to delete user ${email}?`)) {
    return
  }

  try {
    const response = await fetch(`${API_BASE}/users/${encodeURIComponent(email)}`, {
      method: "DELETE",
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to delete user")
    }
    const data = await response.json()
    if (data.success) {
      updateUserTable() // Refresh the table
      showSuccess("User deleted successfully")
    }
  } catch (error) {
    showError(`Failed to delete user: ${error.message}`)
  }
}

// Edit user function
async function editUser(email, clients, access) {
  document.getElementById("userModalTitle").textContent = "Edit User"
  document.getElementById("modalEmail").value = email
  document.getElementById("modalEmail").disabled = true // Disable email editing
  document.getElementById("modalAccess").value = access

  // Always load template folders for the modal
  await loadTemplateFoldersForModal()

  // Clear all checkboxes first
  const checkboxes = document.querySelectorAll('#modalClients input[type="checkbox"]')
  checkboxes.forEach((checkbox) => {
    checkbox.checked = false
  })

  // Check the boxes for user's current clients
  try {
    const parsedClients = JSON.parse(clients)
    parsedClients.forEach((client) => {
      const checkbox = document.getElementById(`client_${client}`)
      if (checkbox) {
        checkbox.checked = true
      }
    })
  } catch (error) {
    console.error("Error parsing clients:", error)
    // If parsing fails, try to split by comma (fallback)
    if (typeof clients === "string") {
      const clientArray = clients.split(",").map((c) => c.trim())
      clientArray.forEach((client) => {
        const checkbox = document.getElementById(`client_${client}`)
        if (checkbox) {
          checkbox.checked = true
        }
      })
    }
  }

  // Store original email for update
  document.getElementById("userForm").dataset.originalEmail = email
  document.getElementById("userModal").classList.remove("hidden")
}

// Show success message
function showSuccess(message) {
  const successDiv = document.createElement("div")
  successDiv.className =
    "fixed top-4 right-4 bg-green-50 border border-green-200 text-green-800 p-4 rounded-xl shadow-large z-50 animate-fade-in"
  successDiv.innerHTML = `
        <div class="flex items-center gap-3">
            <div class="w-8 h-8 bg-green-100 rounded-full flex items-center justify-center">
                <svg class="w-4 h-4 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
                </svg>
            </div>
            <div>
                <p class="font-medium">Success!</p>
                <p class="text-sm">${message}</p>
            </div>
        </div>
    `
  document.body.appendChild(successDiv)
  setTimeout(() => {
    successDiv.remove()
  }, 3000)
}

document.getElementById("addUserButton").addEventListener("click", async () => {
  document.getElementById("userModalTitle").textContent = "Add New User"
  document.getElementById("modalEmail").value = ""
  document.getElementById("modalAccess").value = "user"
  document.getElementById("modalEmail").disabled = false

  // Clear the original email dataset
  delete document.getElementById("userForm").dataset.originalEmail

  // Always load template folders for the modal
  await loadTemplateFoldersForModal()

  // Clear all checkboxes
  const checkboxes = document.querySelectorAll('#modalClients input[type="checkbox"]')
  checkboxes.forEach((checkbox) => {
    checkbox.checked = false
  })

  document.getElementById("userModal").classList.remove("hidden")
})

document.getElementById("cancelUserModal").addEventListener("click", () => {
  document.getElementById("userModal").classList.add("hidden")
})

document.getElementById("userForm").addEventListener("submit", async (e) => {
  e.preventDefault()

  const email = document.getElementById("modalEmail").value
  const access = document.getElementById("modalAccess").value
  const originalEmail = document.getElementById("userForm").dataset.originalEmail

  // Get selected clients from checkboxes
  const selectedClients = []
  const checkboxes = document.querySelectorAll('#modalClients input[type="checkbox"]:checked')
  checkboxes.forEach((checkbox) => {
    selectedClients.push(checkbox.value)
  })

  if (!email || selectedClients.length === 0) {
    showError("Please fill in all fields and select at least one template folder")
    return
  }

  try {
    let response
    if (originalEmail) {
      // Update existing user
      response = await fetch(`${API_BASE}/users/${encodeURIComponent(originalEmail)}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          email,
          clients: selectedClients,
          access,
        }),
        credentials: "include",
      })
    } else {
      // Create new user
      response = await fetch(`${API_BASE}/users`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          email,
          clients: selectedClients,
          access,
        }),
        credentials: "include",
      })
    }

    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || `Failed to ${originalEmail ? "update" : "create"} user`)
    }

    const data = await response.json()
    if (data.success) {
      document.getElementById("userModal").classList.add("hidden")
      // Clear the original email dataset
      delete document.getElementById("userForm").dataset.originalEmail
      updateUserTable() // Refresh the table
      showSuccess(`User ${originalEmail ? "updated" : "created"} successfully`)
    }
  } catch (error) {
    showError(`Failed to ${originalEmail ? "update" : "create"} user: ${error.message}`)
  }
})

// Refresh audit trail
document.getElementById("refreshAuditButton").addEventListener("click", () => {
  loadAuditTrail(1) // Reset to first page
})

// Event listeners for DL Generator
document.getElementById("modeSelect").addEventListener("change", async (e) => {
  const mode = e.target.value
  if (!mode) {
    document.getElementById("selectionSection").classList.add("hidden")
    document.getElementById("transmittalFolderSection").classList.add("hidden")
    document.getElementById("uploadSection").classList.add("hidden")
    document.getElementById("statusDisplay").classList.add("hidden")
    document.getElementById("placeholdersDisplay").classList.add("hidden")
    return
  }
  try {
    const response = await fetch(`${API_BASE}/set_mode`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mode }),
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to set mode")
    }
    const data = await response.json()

    document.getElementById("selectionSection").classList.add("hidden")
    document.getElementById("transmittalFolderSection").classList.add("hidden")
    document.getElementById("uploadSection").classList.add("hidden")
    document.getElementById("placeholdersDisplay").classList.add("hidden")
    document.getElementById("dataPreview").classList.add("hidden")
    clearResultSection()

    // Keep the mode selection
    document.getElementById("modeSelect").value = mode

    if (data.template_status?.transmittal_template) {
      document.getElementById("templateStatusText").textContent = data.template_status.transmittal_template
      document.getElementById("statusDisplay").classList.remove("hidden")
    } else {
      document.getElementById("statusDisplay").classList.add("hidden")
    }

    if (mode === "Transmittal Only") {
      document.getElementById("transmittalFolderSection").classList.remove("hidden")
      await fetchFoldersForTransmittal()
    } else if (mode === "DL w/ Transmittal") {
      document.getElementById("selectionSection").classList.remove("hidden")
      await fetchFolders()
    } else {
      document.getElementById("selectionSection").classList.remove("hidden")
      await fetchFolders()
    }
  } catch (error) {
    showError("Failed to set mode. Please check if the backend server is running.")
    document.getElementById("modeSelect").value = ""
  }
})

async function fetchFoldersForTransmittal() {
  try {
    const response = await fetch(`${API_BASE}/folders`, {
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to fetch folders")
    }
    const folders = await response.json()
    const transmittalFolderSelect = document.getElementById("transmittalFolderSelect")
    transmittalFolderSelect.innerHTML = '<option value="">Select Client Folder</option>'
    folders.forEach((folder) => {
      const option = document.createElement("option")
      option.value = folder
      option.textContent = folder
      transmittalFolderSelect.appendChild(option)
    })
  } catch (error) {
    showError("Failed to fetch folders. Please check FTP configuration.")
  }
}

document.getElementById("transmittalFolderSelect").addEventListener("change", async (e) => {
  const folder = e.target.value
  if (folder) {
    try {
      const response = await fetch(`${API_BASE}/transmittal_placeholders?folder=${encodeURIComponent(folder)}`, {
        method: "GET",
        credentials: "include",
      })
      if (!response.ok) {
        if (response.status === 401) {
          redirectToLogin()
          return
        }
        const errorData = await response.json()
        throw new Error(errorData.detail || "Failed to fetch transmittal placeholders")
      }
      const data = await response.json()

      const placeholdersList = document.getElementById("placeholdersList")
      if (data.message) {
        placeholdersList.innerHTML = ""
        const placeholders = data.placeholders || []
        const filteredPlaceholders = placeholders
          .filter((placeholder) => !placeholder.startsWith("«IMAGE_"))
          .map((placeholder) => placeholder.replace(/«|»/g, ""))

        if (filteredPlaceholders.length > 0) {
          filteredPlaceholders.forEach((placeholder) => {
            const li = document.createElement("li")
            li.className = "flex items-center gap-2 text-sm text-text-secondary"
            li.innerHTML = `
              <div class="w-2 h-2 bg-green-500 rounded-full"></div>
              <code class="bg-green-50 px-2 py-1 rounded text-xs font-mono">${placeholder}</code>
            `
            placeholdersList.appendChild(li)
          })
        } else {
          placeholdersList.innerHTML = '<li class="text-sm text-text-secondary">No valid placeholders found.</li>'
        }

        document.getElementById("placeholdersDisplay").classList.remove("hidden")
        document.getElementById("uploadSection").classList.remove("hidden")
      } else {
        document.getElementById("templateStatusText").textContent = data.detail
        document.getElementById("statusDisplay").classList.remove("hidden")
      }
    } catch (error) {
      console.error("Error processing transmittal placeholders:", error)
      showError("Failed to fetch transmittal placeholders. Please check template configuration.")
    }
  } else {
    document.getElementById("placeholdersDisplay").classList.add("hidden")
    document.getElementById("uploadSection").classList.add("hidden")
  }
})

document.getElementById("outputFormatSelect").addEventListener("change", async (e) => {
  const format = e.target.value
  if (!format) {
    document.querySelector(".card:has(#modeSelect)").classList.add("hidden")
    return
  }

  try {
    const response = await fetch(`${API_BASE}/set_output_format`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ format }),
      credentials: "include",
    })

    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to set output format")
    }

    document.querySelector(".card:has(#modeSelect)").classList.remove("hidden")

    if (format === "print") {
      document.getElementById("printFormatInfo").classList.remove("hidden")
      document.getElementById("zipFormatInfo").classList.add("hidden")
    } else {
      document.getElementById("printFormatInfo").classList.add("hidden")
      document.getElementById("zipFormatInfo").classList.remove("hidden")
    }
  } catch (error) {
    showError("Failed to set output format. Please try again.")
  }
})

document.getElementById("folderSelect").addEventListener("change", (e) => {
  if (e.target.value) {
    fetchDLTypes(e.target.value)
  }
})

document.getElementById("dlTypeSelect").addEventListener("change", (e) => {
  if (e.target.value) {
    fetchTemplates(document.getElementById("folderSelect").value)
  }
})

document.getElementById("templateSelect").addEventListener("change", async (e) => {
  if (e.target.value) {
    const mode = document.getElementById("modeSelect").value
    const folder = document.getElementById("folderSelect").value
    const dlType = document.getElementById("dlTypeSelect").value
    const template = e.target.value

    await fetchPlaceholders(folder, dlType, template)

    if (mode === "DL w/ Transmittal") {
      document.getElementById("placeholdersLoadingOverlay").classList.remove("hidden")
      try {
        const response = await fetch(`${API_BASE}/transmittal_placeholders`, {
          method: "GET",
          credentials: "include",
        })

        if (response.ok) {
          const data = await response.json()
          const placeholdersList = document.getElementById("placeholdersList")

          const separator = document.createElement("li")
          separator.className = "py-2 border-t border-border-light mt-2 pt-2"
          separator.innerHTML = `
            <div class="flex items-center gap-2">
              <div class="w-4 h-4 bg-purple-500 rounded-full"></div>
              <span class="font-medium text-purple-700">Transmittal Placeholders</span>
            </div>
          `
          placeholdersList.appendChild(separator)

          const transmittalPlaceholders = data.placeholders || []
          const filteredTransmittalPlaceholders = transmittalPlaceholders
            .filter((placeholder) => !placeholder.startsWith("«IMAGE_"))
            .map((placeholder) => placeholder.replace(/«|»/g, ""))

          if (filteredTransmittalPlaceholders.length > 0) {
            filteredTransmittalPlaceholders.forEach((placeholder) => {
              const li = document.createElement("li")
              li.className = "flex items-center gap-2 text-sm text-purple-600"
              li.innerHTML = `
                <div class="w-2 h-2 bg-purple-500 rounded-full"></div>
                <code class="bg-purple-50 px-2 py-1 rounded text-xs font-mono">${placeholder}</code>
              `
              placeholdersList.appendChild(li)
            })
          } else {
            const li = document.createElement("li")
            li.className = "text-sm text-purple-600"
            li.textContent = "No transmittal placeholders found."
            placeholdersList.appendChild(li)
          }
        }
      } catch (error) {
        console.error("Error fetching transmittal placeholders:", error)
      } finally {
        document.getElementById("placeholdersLoadingOverlay").classList.add("hidden")
      }
    }
  }
})

document.getElementById("excelUpload").addEventListener("change", async (e) => {
  const file = e.target.files[0]
  if (file) {
    document.getElementById("excelLoadingOverlay").classList.remove("hidden")
    const formData = new FormData()
    formData.append("file", file)
    try {
      const response = await fetch(`${API_BASE}/upload_excel`, {
        method: "POST",
        body: formData,
        credentials: "include",
      })
      const data = await response.json()
      document.getElementById("excelLoadingOverlay").classList.add("hidden")

      if (!response.ok) {
        showError(data.detail || "An error occurred while uploading the Excel file.")
        document.getElementById("excelUpload").value = ""
        return
      }

      if (data && data.data && Array.isArray(data.data) && data.data.length > 0) {
        const totalRows = data.data.length
        const previewRows = data.data.slice(0, 10) // Show only first 10 rows

        // Create row count display
        const rowCountDisplay = document.createElement("div")
        rowCountDisplay.className = "mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg"
        rowCountDisplay.innerHTML = `
          <div class="flex items-center gap-3">
            <div class="w-8 h-8 bg-blue-100 rounded-lg flex items-center justify-center">
              <svg class="w-4 h-4 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
              </svg>
            </div>
            <div>
              <p class="font-semibold text-blue-800">Excel File Loaded Successfully</p>
              <p class="text-sm text-blue-600">
                <span class="font-medium">${totalRows} data rows</span> found 
                ${totalRows > 10 ? `(showing first 10 rows in preview)` : ""}
              </p>
            </div>
          </div>
        `

        const tableContainer = document.getElementById("dataTable")
        tableContainer.innerHTML = ""
        tableContainer.className = "space-y-4"

        // Add row count display
        tableContainer.appendChild(rowCountDisplay)

        // Create table container
        const tableWrapper = document.createElement("div")
        tableWrapper.className = "max-h-[400px] overflow-auto rounded-lg border border-border-light"

        const table = document.createElement("table")
        table.className = "min-w-full text-sm"

        // Create header
        const thead = document.createElement("thead")
        const headerRow = document.createElement("tr")
        headerRow.className = "border-b border-border-light bg-white"
        Object.keys(data.data[0]).forEach((key) => {
          const th = document.createElement("th")
          th.className = "sticky top-0 bg-white z-10 text-left py-3 px-4 font-medium text-text-secondary"
          th.textContent = key
          headerRow.appendChild(th)
        })
        thead.appendChild(headerRow)
        table.appendChild(thead)

        // Create body with limited rows
        const tbody = document.createElement("tbody")
        tbody.className = "divide-y divide-border-light"
        previewRows.forEach((row, index) => {
          const tr = document.createElement("tr")
          tr.className = `hover:bg-surface-hover transition-colors ${index % 2 === 0 ? "bg-gray-50" : "bg-white"}`
          Object.values(row).forEach((value) => {
            const td = document.createElement("td")
            td.className = "px-4 py-3 text-text-secondary"
            td.textContent = value
            tr.appendChild(td)
          })
          tbody.appendChild(tr)
        })

        // Add "more rows" indicator if there are more than 10 rows
        if (totalRows > 10) {
          const moreRowsIndicator = document.createElement("tr")
          moreRowsIndicator.className = "bg-gray-100"
          const td = document.createElement("td")
          td.colSpan = Object.keys(data.data[0]).length
          td.className = "px-4 py-3 text-center text-gray-500 italic"
          td.textContent = `... and ${totalRows - 10} more rows`
          moreRowsIndicator.appendChild(td)
          tbody.appendChild(moreRowsIndicator)
        }

        table.appendChild(tbody)
        tableWrapper.appendChild(table)
        tableContainer.appendChild(tableWrapper)

        document.getElementById("dataPreview").classList.remove("hidden")
      } else {
        showError("Uploaded Excel file is empty or has an invalid structure.")
        document.getElementById("dataPreview").classList.add("hidden")
      }
    } catch (error) {
      document.getElementById("excelLoadingOverlay").classList.add("hidden")
      console.error("Excel upload error:", error)
      showError(`Failed to upload Excel file: ${error.message}. Please check file format and network connection.`)
      document.getElementById("excelUpload").value = ""
    }
  }
})

async function loadAvailablePrinters() {
  try {
    const response = await fetch(`${API_BASE}/printers`, {
      credentials: "include",
    })
    if (response.ok) {
      const data = await response.json()
      return data.printers
    } else {
      console.error("Failed to load printers:", response.status)
      return []
    }
  } catch (error) {
    console.error("Failed to load printers:", error)
    return []
  }
}

function createPrintControls(areas, printers) {
  const printContainer = document.createElement("div")
  printContainer.className = "mt-6 p-4 bg-gray-50 rounded-lg border border-gray-200"

  printContainer.innerHTML = `
    <h4 class="text-lg font-semibold text-gray-800 mb-4">🖨️ Print Documents</h4>
    <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
      <div>
        <label class="block text-sm font-medium text-gray-600 mb-2">Select Area</label>
        <select id="printAreaSelect" class="w-full p-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
          <option value="">Choose area to print</option>
          ${areas.map((area) => `<option value="${area}">${area}</option>`).join("")}
        </select>
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-600 mb-2">Select Printer</label>
        <select id="printerSelect" class="w-full p-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500">
          <option value="">Default Printer</option>
          ${printers
            .map(
              (printer) => `
            <option value="${printer.name}" ${printer.is_default ? "selected" : ""}>
              ${printer.name}${printer.is_default ? " (Default)" : ""}
            </option>
          `,
            )
            .join("")}
        </select>
      </div>
      <div class="flex items-end">
        <button id="printButton" class="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 px-4 rounded-lg font-medium flex items-center justify-center gap-2 transition-colors">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path>
          </svg>
          Print Selected Area
        </button>
      </div>
    </div>
    <div id="printStatus" class="hidden mt-3 p-3 rounded-lg"></div>
  `

  // Add print functionality
  const printButton = printContainer.querySelector("#printButton")
  const printAreaSelect = printContainer.querySelector("#printAreaSelect")
  const printerSelect = printContainer.querySelector("#printerSelect")
  const printStatus = printContainer.querySelector("#printStatus")

  printButton.addEventListener("click", async () => {
    const selectedArea = printAreaSelect.value
    const selectedPrinter = printerSelect.value

    if (!selectedArea) {
      showPrintStatus("Please select an area to print", "error")
      return
    }

    try {
      printButton.disabled = true
      printButton.innerHTML = `
        <div class="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
        Printing...
      `

      let printUrl = `${API_BASE}/print_files/${selectedArea}`
      if (selectedPrinter) {
        printUrl += `?printer=${encodeURIComponent(selectedPrinter)}`
      }

      const response = await fetch(printUrl, {
        credentials: "include",
      })

      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.detail || "Failed to print files")
      }

      const result = await response.json()
      showPrintStatus(result.message, "success")
    } catch (error) {
      showPrintStatus(`Failed to print: ${error.message}`, "error")
    } finally {
      printButton.disabled = false
      printButton.innerHTML = `
        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path>
        </svg>
        Print Selected Area
      `
    }
  })

  function showPrintStatus(message, type) {
    const bgColor =
      type === "success" ? "bg-green-50 border-green-200 text-green-800" : "bg-red-50 border-red-200 text-red-800"
    const icon = type === "success" ? "✅" : "❌"

    printStatus.className = `mt-3 p-3 rounded-lg border ${bgColor}`
    printStatus.innerHTML = `${icon} ${message}`
    printStatus.classList.remove("hidden")

    setTimeout(() => {
      printStatus.classList.add("hidden")
    }, 5000)
  }

  return printContainer
}

// Cleanup button handler
document.getElementById("cleanupButton").addEventListener("click", async () => {
  try {
    const response = await fetch(`${API_BASE}/cleanup`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "include",
    })
    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to cleanup files")
    }
    const data = await response.json()
    if (data.success) {
      document.getElementById("resultSection").classList.add("hidden")
      document.getElementById("progressSection").classList.add("hidden")
      resetUI()
    } else {
      showError(data.detail || "Failed to cleanup files")
    }
  } catch (error) {
    showError("Failed to cleanup files. Please check backend server.")
  }
})

// User Management Functions
const editUserIndex = null

async function updateUserTable() {
  try {
    // Show loading state
    const tbody = document.getElementById("userTableBody")
    const loadingRow = document.getElementById("loadingState")
    if (loadingRow) {
      loadingRow.classList.remove("hidden")
    }

    const response = await fetch(`${API_BASE}/users`, {
      credentials: "include",
    })

    if (!response.ok) {
      if (response.status === 401) {
        redirectToLogin()
        return
      }
      if (response.status === 403) {
        showError("Admin access required for user management.")
        return
      }
      throw new Error("Failed to fetch users")
    }

    const users = await response.json()
    const userTableBody = document.getElementById("userTableBody")
    userTableBody.innerHTML = "" // Clear loading state and prior content

    // Update stats
    updateUserStats(users)

    if (users.length === 0) {
      userTableBody.innerHTML = `
                <tr>
                    <td colspan="4" class="px-8 py-16 text-center text-gray-500">
                        <div class="flex flex-col items-center gap-4">
                            <div class="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center">
                                <svg class="w-8 h-8 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197m13.5-9a2.5 2.5 0 11-5 0 2.5 2.5 0 015 0z"></path>
                                </svg>
                            </div>
                            <div>
                                <p class="text-lg font-medium text-gray-900">No users found</p>
                                <p class="text-sm text-gray-500 mt-1">Get started by adding your first user</p>
                            </div>
                            <button onclick="document.getElementById('addUserButton').click()" 
                                    class="inline-flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white py-2 px-4 rounded-lg font-medium transition-colors">
                                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"></path>
                                </svg>
                                Add First User
                            </button>
                        </div>
                    </td>
                </tr>
            `
      return
    }

    users.forEach((user, index) => {
      const row = document.createElement("tr")
      row.className = "hover:bg-blue-50 transition-all duration-200 ease-in-out group"
      const clientsDisplay = Array.isArray(user.clients) ? user.clients.join(", ") : user.clients
      row.innerHTML = `
                <td class="px-8 py-6">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 bg-gradient-to-br from-blue-500 to-purple-600 rounded-full flex items-center justify-center">
                            <span class="text-white font-semibold text-sm">${user.email.charAt(0).toUpperCase()}</span>
                        </div>
                        <div>
                            <p class="font-semibold text-gray-900">${user.email}</p>
                            <p class="text-sm text-gray-500">User ID: ${index + 1}</p>
                        </div>
                    </div>
                </td>
                <td class="px-8 py-6">
                    <div class="flex flex-wrap gap-1">
                        ${
                          Array.isArray(user.clients)
                            ? user.clients
                                .map(
                                  (client) =>
                                    `<span class="inline-flex items-center px-3 py-1 text-xs font-medium bg-blue-100 text-blue-800 rounded-full border border-blue-200">${client}</span>`,
                                )
                                .join("")
                            : `<span class="inline-flex items-center px-3 py-1 text-xs font-medium bg-blue-100 text-blue-800 rounded-full border border-blue-200">${user.clients}</span>`
                        }
                    </div>
                </td>
                <td class="px-8 py-6">
                    <div class="flex items-center gap-2">
                        <div class="w-3 h-3 rounded-full ${user.access === "admin" ? "bg-green-400" : "bg-blue-400"}"></div>
                        <span class="inline-flex items-center px-3 py-1 text-sm font-medium rounded-full ${
                          user.access === "admin"
                            ? "bg-green-100 text-green-800 border border-green-200"
                            : "bg-gray-100 text-gray-800 border border-gray-200"
                        }">
                            ${user.access === "admin" ? "Administrator" : "User"}
                        </span>
                    </div>
                </td>
                <td class="px-8 py-6">
                    <div class="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity duration-200">
                        <button onclick="editUser('${user.email}', decodeURIComponent('${encodeURIComponent(JSON.stringify(user.clients))}'), '${user.access}')"
                            class="inline-flex items-center gap-1 px-3 py-2 text-sm font-medium text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-all duration-200" 
                            title="Edit User">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
                            </svg>
                            Edit
                        </button>
                        <button onclick="deleteUser('${user.email}')" 
                            class="inline-flex items-center gap-1 px-3 py-2 text-sm font-medium text-red-600 hover:text-red-800 hover:bg-red-50 rounded-lg transition-all duration-200" 
                            title="Delete User">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5-4h4a1 1 0 011 1v1H9V4a1 1 0 011-1z"></path>
                            </svg>
                            Delete
                        </button>
                    </div>
                </td>
            `
      userTableBody.appendChild(row)
    })

    // Setup search and filter functionality
    setupUserFilters(users)
  } catch (error) {
    const userTableBody = document.getElementById("userTableBody")
    userTableBody.innerHTML = `
            <tr>
                <td colspan="4" class="px-8 py-12 text-center text-red-500">
                    <div class="flex flex-col items-center gap-3">
                        <svg class="w-12 h-12 text-red-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                        </svg>
                        <p class="text-lg font-medium">Failed to load users</p>
                        <p class="text-sm">Please try refreshing the page</p>
                    </div>
                </td>
            </tr>
        `
    showError("Failed to load users.")
  }
}

function updateUserStats(users) {
  const totalUsers = users.length
  const adminUsers = users.filter((user) => user.access === "admin").length
  const regularUsers = users.filter((user) => user.access === "user").length

  // Update stat cards
  const totalUsersElement = document.getElementById("totalUsersCount")
  const adminUsersElement = document.getElementById("adminUsersCount")
  const regularUsersElement = document.getElementById("regularUsersCount")

  if (totalUsersElement) totalUsersElement.textContent = totalUsers
  if (adminUsersElement) adminUsersElement.textContent = adminUsers
  if (regularUsersElement) regularUsersElement.textContent = regularUsers
}

function setupUserFilters(allUsers) {
  const searchInput = document.getElementById("userSearchInput")
  const accessFilter = document.getElementById("accessFilter")
  const clearFilters = document.getElementById("clearFilters")

  function filterUsers() {
    const searchTerm = searchInput?.value.toLowerCase() || ""
    const accessLevel = accessFilter?.value || ""

    const filteredUsers = allUsers.filter((user) => {
      const matchesSearch = user.email.toLowerCase().includes(searchTerm)
      const matchesAccess = !accessLevel || user.access === accessLevel
      return matchesSearch && matchesAccess
    })

    renderFilteredUsers(filteredUsers)
  }

  function renderFilteredUsers(users) {
    const userTableBody = document.getElementById("userTableBody")
    userTableBody.innerHTML = ""

    if (users.length === 0) {
      userTableBody.innerHTML = `
                <tr>
                    <td colspan="4" class="px-8 py-12 text-center text-gray-500">
                        <div class="flex flex-col items-center gap-3">
                            <svg class="w-12 h-12 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path>
                            </svg>
                            <p class="text-lg font-medium">No users match your filters</p>
                            <p class="text-sm">Try adjusting your search criteria</p>
                        </div>
                    </td>
                </tr>
            `
      return
    }

    users.forEach((user, index) => {
      const row = document.createElement("tr")
      row.className = "hover:bg-blue-50 transition-all duration-200 ease-in-out group"
      row.innerHTML = `
                <td class="px-8 py-6">
                    <div class="flex items-center gap-3">
                        <div class="w-10 h-10 bg-gradient-to-br from-blue-500 to-purple-600 rounded-full flex items-center justify-center">
                            <span class="text-white font-semibold text-sm">${user.email.charAt(0).toUpperCase()}</span>
                        </div>
                        <div>
                            <p class="font-semibold text-gray-900">${user.email}</p>
                            <p class="text-sm text-gray-500">User ID: ${index + 1}</p>
                        </div>
                    </div>
                </td>
                <td class="px-8 py-6">
                    <div class="flex flex-wrap gap-1">
                        ${
                          Array.isArray(user.clients)
                            ? user.clients
                                .map(
                                  (client) =>
                                    `<span class="inline-flex items-center px-3 py-1 text-xs font-medium bg-blue-100 text-blue-800 rounded-full border border-blue-200">${client}</span>`,
                                )
                                .join("")
                            : `<span class="inline-flex items-center px-3 py-1 text-xs font-medium bg-blue-100 text-blue-800 rounded-full border border-blue-200">${user.clients}</span>`
                        }
                    </div>
                </td>
                <td class="px-8 py-6">
                    <div class="flex items-center gap-2">
                        <div class="w-3 h-3 rounded-full ${user.access === "admin" ? "bg-green-400" : "bg-blue-400"}"></div>
                        <span class="inline-flex items-center px-3 py-1 text-sm font-medium rounded-full ${
                          user.access === "admin"
                            ? "bg-green-100 text-green-800 border border-green-200"
                            : "bg-gray-100 text-gray-800 border border-gray-200"
                        }">
                            ${user.access === "admin" ? "Administrator" : "Standard User"}
                        </span>
                    </div>
                </td>
                <td class="px-8 py-6">
                    <div class="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity duration-200">
                        <button onclick="editUser('${user.email}', decodeURIComponent('${encodeURIComponent(JSON.stringify(user.clients))}'), '${user.access}')"
                            class="inline-flex items-center gap-1 px-3 py-2 text-sm font-medium text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-all duration-200" 
                            title="Edit User">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
                            </svg>
                            Edit
                        </button>
                        <button onclick="deleteUser('${user.email}')" 
                            class="inline-flex items-center gap-1 px-3 py-2 text-sm font-medium text-red-600 hover:text-red-800 hover:bg-red-50 rounded-lg transition-all duration-200" 
                            title="Delete User">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5-4h4a1 1 0 011 1v1H9V4a1 1 0 011-1z"></path>
                            </svg>
                            Delete
                        </button>
                    </div>
                </td>
            `
      userTableBody.appendChild(row)
    })
  }

  // Add event listeners
  if (searchInput) {
    searchInput.addEventListener("input", filterUsers)
  }
  if (accessFilter) {
    accessFilter.addEventListener("change", filterUsers)
  }
  if (clearFilters) {
    clearFilters.addEventListener("click", () => {
      if (searchInput) searchInput.value = ""
      if (accessFilter) accessFilter.value = ""
      renderFilteredUsers(allUsers)
    })
  }
}

// Enable drag and drop for Excel file
const dropZone = document.querySelector("label[for='excelUpload']")
if (dropZone) {
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault()
    if (!isProcessing) {
      dropZone.classList.add("border-primary", "bg-primary", "bg-opacity-5")
    }
  })
  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("border-primary", "bg-primary", "bg-opacity-5")
  })
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault()
    dropZone.classList.remove("border-primary", "bg-primary", "bg-opacity-5")
    if (!isProcessing) {
      document.getElementById("excelUpload").files = e.dataTransfer.files
      document.getElementById("excelUpload").dispatchEvent(new Event("change"))
    }
  })
}

// Make functions globally available
window.deleteUser = deleteUser
window.editUser = editUser
window.showAuditDetails = showAuditDetails
