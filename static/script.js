let currentUser = null
const API_BASE = "http://localhost:5000/api"
let isProcessing = false // Track processing state

// Initialize UI and check session
document.addEventListener("DOMContentLoaded", async () => {
  showSection("dlGeneratorSection")
  // Hide mode selection initially until output format is selected
  document.querySelector(".card:has(#modeSelect)").classList.add("hidden")
  setupMobileMenu()
  setupFolderSelectionButtons()
  const urlParams = new URLSearchParams(window.location.search)
  if (urlParams.get("code")) {
    await handleLarkCallback(urlParams.get("code"))
  } else {
    await checkSession()
  }
})

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

  // Close mobile menu when clicking nav items
  const navItems = document.querySelectorAll(".nav-item")
  navItems.forEach((item) => {
    item.addEventListener("click", () => {
      sidebar.classList.remove("open")
      mobileMenuOverlay.classList.add("hidden")
    })
  })
}

// Setup folder selection buttons
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

// Function to disable/enable all input fields
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

  // Also disable the upload label
  const uploadLabel = document.getElementById("uploadLabel")
  if (uploadLabel) {
    if (disabled) {
      uploadLabel.classList.add("disabled-input")
    } else {
      uploadLabel.classList.remove("disabled-input")
    }
  }
}

// Check existing session
async function checkSession() {
  try {
    const response = await fetch(`${API_BASE}/check_sessions`, {
      method: "GET",
      credentials: "include",
    })
    console.log("sample",response)
    if (response.ok) {
      const data = await response.json()
      console.log(data)

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

        // Set user avatar
        setUserAvatar(currentUser.userInfo)
        console.log(currentUser)

        // Show/hide navigation based on access level
        updateNavigationAccess()

        fetchFolders()
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

// Set user avatar
function setUserAvatar(userInfo) {
  console.log("User info for avatar:", userInfo)
  const userAvatar = document.getElementById("userAvatar")
  const userAvatarFallback = document.getElementById("userAvatarFallback")

  if (userInfo) {
    userAvatar.src = userInfo
    userAvatar.classList.remove("hidden")
    userAvatarFallback.classList.add("hidden")
    console.log("Avatar set to:", userInfo)
  } else {
    userAvatar.classList.add("hidden")
    userAvatarFallback.classList.remove("hidden")
    console.log("No avatar URL found, using fallback")
  }
}

// Update navigation access based on user role
function updateNavigationAccess() {
  const userManagementMenu = document.getElementById("userManagementMenu")
  const auditTrailMenu = document.getElementById("auditTrailMenu")

  if (currentUser && currentUser.access === "admin") {
    // Admin can see all menu items
    userManagementMenu.style.display = "flex"
    auditTrailMenu.style.display = "flex"
  } else {
    // User can only see DL Generator
    userManagementMenu.style.display = "none"
    auditTrailMenu.style.display = "none"
  }
}

// Handle Lark callback
let isProcessingCallback = false;

async function handleLarkCallback(code) {
    if (isProcessingCallback) {
        console.log("Already processing callback, ignoring duplicate");
        return;
    }
    isProcessingCallback = true;
    console.log(`Handling Lark callback with code: ${code}`);
    try {
        const response = await fetch(`${API_BASE}/lark_callback?code=${code}`, {
            method: "GET",
            credentials: "include",
        });
        const data = await response.json();
        console.log("Lark callback response:", data);
        if (data.success) {
            currentUser = { username: data.username, role: data.role };
            document.getElementById("loginModal").classList.add("hidden");
            document.getElementById("mainContent").classList.remove("hidden");
            document.getElementById("userDisplay").textContent = `${data.username} (${data.role})`;
            window.history.replaceState({}, document.title, "/");
            await checkSession();
        } else {
            document.getElementById("loginError").classList.remove("hidden");
            document.getElementById("loginError").textContent = data.detail || "Authentication failed.";
        }
    } catch (error) {
        console.error("Error in handleLarkCallback:", error);
        document.getElementById("loginError").classList.remove("hidden");
        document.getElementById("loginError").textContent = "Authentication failed. Please try again.";
    } finally {
        isProcessingCallback = false;
    }
}
// Initiate Lark login
document.getElementById("loginButton").addEventListener("click", () => {
  window.location.href = `${API_BASE}/login`
})

// Logout handling
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
      window.history.replaceState({}, document.title, "/")
    } else {
      showError(data.detail || "Logout failed.")
    }
  } catch (error) {
    showError("Logout failed. Please check if the backend server is running.")
  }
})

// Show login modal
function showLoginModal() {
  document.getElementById("mainContent").classList.add("hidden")
  document.getElementById("loginModal").classList.remove("hidden")
  document.getElementById("loginError").classList.add("hidden")
}

// Redirect to login on session expiration
function redirectToLogin() {
  currentUser = null
  showLoginModal()
  resetUI()
  showError("Session expired. Please log in again.")
}

// Update page title and subtitle
function updatePageTitle(title, subtitle) {
  document.getElementById("pageTitle").textContent = title
  document.getElementById("pageSubtitle").textContent = subtitle
}

// Sidebar navigation
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

// Show specific section
function showSection(sectionId) {
  document.getElementById("dlGeneratorSection").classList.add("hidden")
  document.getElementById("userManagementSection").classList.add("hidden")
  document.getElementById("auditTrailSection").classList.add("hidden")
  document.getElementById(sectionId).classList.remove("hidden")

  // Update navigation active states
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

// Reset UI - FIXED: Preserve output format selection and prevent field resets
function resetUI() {
  document.getElementById("outputFormatSelect").value = ""
  document.getElementById("modeSelect").value = ""
  document.getElementById("transmittalFolderSelect").value = ""
  document.getElementById("modeCard").classList.add("hidden")
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

  // FIXED: Clear result section properly to prevent button duplication
  clearResultSection()

  document.getElementById("progressBar").style.width = "0%"
  document.getElementById("progressText").textContent = ""
  document.getElementById("excelUpload").value = ""
  document.getElementById("dataTable").innerHTML = ""
  document.getElementById("statusDisplay").classList.add("hidden")
  document.getElementById("templateCombinedAlert").classList.add("hidden")
  document.getElementById("excelLoadingOverlay").classList.add("hidden")

  // Reset processing state and re-enable inputs
  isProcessing = false
  toggleInputFields(false)
}

// FIXED: New function to properly clear result section
function clearResultSection() {
  const resultSection = document.getElementById("resultSection")
  resultSection.classList.add("hidden")

  // Remove any dynamically added print controls
  const printControls = resultSection.querySelectorAll(".mt-4")
  printControls.forEach((control) => {
    if (control.querySelector("#printAreaSelect") || control.querySelector("#printerSelect")) {
      control.remove()
    }
  })

  // Reset buttons to hidden state
  document.getElementById("downloadButton").classList.add("hidden")
  document.getElementById("cleanupButton").classList.add("hidden")
}

// Show error
function showError(message) {
  const errorDiv = document.getElementById("errorMessage");
  if (!errorDiv) {
    console.error("Error: #errorMessage div not found in the DOM");
    return;
  }

  const errorSpan = errorDiv.querySelector("span");
  if (!errorSpan) {
    console.error("Error: <span> element not found inside #errorMessage");
    return;
  }

  errorSpan.textContent = message || "An unexpected error occurred.";
  errorDiv.classList.remove("hidden");
}
// Show template combined alert
function showTemplateCombinedAlert() {
  document.getElementById("templateCombinedAlert").classList.remove("hidden")
}

// Fetch folders (now user-specific)
async function fetchFolders() {
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
    const folderSelect = document.getElementById("folderSelect")
    folderSelect.innerHTML = '<option value="">Select Folder</option>'
    folders.forEach((folder) => {
      const option = document.createElement("option")
      option.value = folder
      option.textContent = folder
      folderSelect.appendChild(option)
    })

    // Always load template folders for modal if user is admin
    if (currentUser && currentUser.access === "admin") {
      await loadTemplateFoldersForModal()
    }
  } catch (error) {
    showError("Failed to fetch folders. Please check FTP configuration.")
  }
}

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
      const errorData = await response.json()
      throw new Error(errorData.detail || "Failed to fetch DL types")
    }
    const dlTypes = await response.json()
    const dlTypeSelect = document.getElementById("dlTypeSelect")
    dlTypeSelect.innerHTML = '<option value="">Select DL Type</option>'
    dlTypes.forEach((type) => {
      const option = document.createElement("option")
      option.value = type
      option.textContent = type
      dlTypeSelect.appendChild(option)
    })
  } catch (error) {
    showError("Failed to fetch DL types. Please check Google Sheets configuration.")
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

    const placeholdersList = document.getElementById("placeholdersList");
    if (data.message) {
      placeholdersList.innerHTML = "";
      const placeholders = data.placeholders || [];
      // Filter out placeholders starting with "IMAGE_" and clean «»
      const filteredPlaceholders = placeholders
        .filter((placeholder) => !placeholder.startsWith("«IMAGE_"))
        .map((placeholder) => placeholder.replace(/«|»/g, ""));

      if (filteredPlaceholders.length > 0) {
        filteredPlaceholders.forEach((placeholder) => {
          const li = document.createElement("li");
          li.className = "flex items-center gap-2 text-sm text-text-secondary";
          li.innerHTML = `
            <div class="w-2 h-2 bg-accent rounded-full"></div>
            <code class="bg-background px-2 py-1 rounded text-xs font-mono">${placeholder}</code>
          `;
          placeholdersList.appendChild(li);
        });
      } else {
        placeholdersList.innerHTML = '<li class="text-sm text-text-secondary">No valid placeholders found.</li>';
      }

      document.getElementById("placeholdersDisplay").classList.remove("hidden");
      document.getElementById("uploadSection").classList.remove("hidden");

      // Check if template is already combined and show alert
      if (data.template_combined === true) {
        showTemplateCombinedAlert();
      }
    } else {
      document.getElementById("templatecontentStatusText").textContent = data.detail;
      document.getElementById("templatecontentStatus").classList.remove("hidden");
    }
  } catch (error) {
    console.error("Error processing placeholders:", error);
    showError("Failed to fetch placeholders. Please check template configuration.");
  }
}

async function fetchTransmittalPlaceholders() {
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
  // Create a temporary success message element
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
// Replace the existing modeSelect event listener with this updated version
document.getElementById("modeSelect").addEventListener("change", async (e) => {
  const mode = e.target.value
  if (!mode) {
    // FIXED: Don't reset everything when mode is cleared, just clear mode-specific sections
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

    // FIXED: Don't reset UI completely, just clear mode-specific sections
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
      // For Transmittal Only mode, show folder selection first
      document.getElementById("transmittalFolderSection").classList.remove("hidden")
      await fetchFoldersForTransmittal()
    } else if (mode === "DL w/ Transmittal") {
      // For DL w/ Transmittal, show the selection section first
      document.getElementById("selectionSection").classList.remove("hidden")
      await fetchFolders()
    } else {
      // For DL Only mode
      document.getElementById("selectionSection").classList.remove("hidden")
      await fetchFolders()
    }
  } catch (error) {
    showError("Failed to set mode. Please check if the backend server is running.")
    document.getElementById("modeSelect").value = ""
  }
})

// Add new function to fetch folders for transmittal folder selection
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

// Add event listener for transmittal folder selection
document.getElementById("transmittalFolderSelect").addEventListener("change", async (e) => {
  const folder = e.target.value
  if (folder) {
    // Fetch transmittal placeholders with the selected folder
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
        // Filter out placeholders starting with "IMAGE_" and clean «»
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
    // Hide placeholders and upload sections if no folder is selected
    document.getElementById("placeholdersDisplay").classList.add("hidden")
    document.getElementById("uploadSection").classList.add("hidden")
  }
})

// FIXED: Output Format Selection - preserve selection and don't reset when mode changes
document.getElementById("outputFormatSelect").addEventListener("change", async (e) => {
  const format = e.target.value
  if (!format) {
    // Hide mode selection if no format is selected
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

    // Show mode selection after output format is selected
    document.querySelector(".card:has(#modeSelect)").classList.remove("hidden")

    // Update UI based on selected format
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

// document.getElementById("templateSelect").addEventListener("change", (e) => {
//   if (e.target.value) {
//     fetchPlaceholders(
//       document.getElementById("folderSelect").value,
//       document.getElementById("dlTypeSelect").value,
//       e.target.value,
//     )
//   }
// })

// Replace the existing modeSelect event listener with this updated version
// document.getElementById("modeSelect").addEventListener("change", async (e) => {
//   const mode = e.target.value
//   if (!mode) {
//     // FIXED: Don't reset everything when mode is cleared, just clear mode-specific sections
//     document.getElementById("selectionSection").classList.add("hidden")
//     document.getElementById("uploadSection").classList.add("hidden")
//     document.getElementById("statusDisplay").classList.add("hidden")
//     document.getElementById("placeholdersDisplay").classList.add("hidden")
//     return
//   }
//   try {
//     const response = await fetch(`${API_BASE}/set_mode`, {
//       method: "POST",
//       headers: { "Content-Type": "application/json" },
//       body: JSON.stringify({ mode }),
//       credentials: "include",
//     })
//     if (!response.ok) {
//       if (response.status === 401) {
//         redirectToLogin()
//         return
//       }
//       const errorData = await response.json()
//       throw new Error(errorData.detail || "Failed to set mode")
//     }
//     const data = await response.json()

//     // FIXED: Don't reset UI completely, just clear mode-specific sections
//     document.getElementById("selectionSection").classList.add("hidden")
//     document.getElementById("uploadSection").classList.add("hidden")
//     document.getElementById("placeholdersDisplay").classList.add("hidden")
//     document.getElementById("dataPreview").classList.add("hidden")
//     clearResultSection()

//     // Keep the mode selection
//     document.getElementById("modeSelect").value = mode

//     if (data.template_status?.transmittal_template) {
//       document.getElementById("templateStatusText").textContent = data.template_status.transmittal_template
//       document.getElementById("statusDisplay").classList.remove("hidden")
//     } else {
//       document.getElementById("statusDisplay").classList.add("hidden")
//     }

//     if (mode === "Transmittal Only") {
//       // For Transmittal Only mode, fetch transmittal placeholders directly
//       await fetchTransmittalPlaceholders()
//       document.getElementById("uploadSection").classList.remove("hidden")
//     } else if (mode === "DL w/ Transmittal") {
//       // For DL w/ Transmittal, show the selection section first
//       document.getElementById("selectionSection").classList.remove("hidden")
//       await fetchFolders()
//     } else {
//       // For DL Only mode
//       document.getElementById("selectionSection").classList.remove("hidden")
//       await fetchFolders()
//     }
//   } catch (error) {
//     showError("Failed to set mode. Please check if the backend server is running.")
//     document.getElementById("modeSelect").value = ""
//   }
// })

// Update the template selection handler to fetch transmittal placeholders when in DL w/ Transmittal mode
document.getElementById("templateSelect").addEventListener("change", async (e) => {
  if (e.target.value) {
    const mode = document.getElementById("modeSelect").value
    const folder = document.getElementById("folderSelect").value
    const dlType = document.getElementById("dlTypeSelect").value
    const template = e.target.value

    // First fetch regular template placeholders
    await fetchPlaceholders(folder, dlType, template)

    // If in DL w/ Transmittal mode, also fetch transmittal placeholders and combine them
    if (mode === "DL w/ Transmittal") {
      try {
        const response = await fetch(`${API_BASE}/transmittal_placeholders`, {
          method: "GET",
          credentials: "include",
        })

        if (response.ok) {
          const data = await response.json()
          const placeholdersList = document.getElementById("placeholdersList")

          // Add a separator for transmittal placeholders
          const separator = document.createElement("li")
          separator.className = "py-2 border-t border-border-light mt-2 pt-2"
          separator.innerHTML = `
            <div class="flex items-center gap-2">
              <div class="w-4 h-4 bg-purple-500 rounded-full"></div>
              <span class="font-medium text-purple-700">Transmittal Placeholders</span>
            </div>
          `
          placeholdersList.appendChild(separator)

          // Add transmittal placeholders
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
      }
    }
  }
})

document.getElementById("excelUpload").addEventListener("change", async (e) => {
  const file = e.target.files[0]
  if (file) {
    // Show loading overlay
    document.getElementById("excelLoadingOverlay").classList.remove("hidden")

    const formData = new FormData()
    formData.append("file", file)
    try {
      const response = await fetch(`${API_BASE}/upload_excel`, {
        method: "POST",
        body: formData,
      })
      const data = await response.json()
      
      // Hide loading overlay
      document.getElementById("excelLoadingOverlay").classList.add("hidden")

      if (!response.ok) {
        // Display error message from server
        console.log("1",  data.detail)
        showError(data.detail || "An error occurred while uploading the file.");
        return;
      }

      const tableContainer = document.getElementById("dataTable")
      tableContainer.innerHTML = ""
      tableContainer.className = "max-h-[500px] overflow-auto rounded-lg border border-border-light"

      // Create table
      const table = document.createElement("table")
      table.className = "min-w-full text-sm"

      // Create thead with sticky header
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

      // Create tbody
      const tbody = document.createElement("tbody")
      tbody.className = "divide-y divide-border-light"

      data.data.forEach((row, index) => {
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

      table.appendChild(tbody)
      tableContainer.appendChild(table)

      // Show data preview section
      document.getElementById("dataPreview").classList.remove("hidden")
    } catch (error) {
      // Hide loading overlay on error
      document.getElementById("excelLoadingOverlay").classList.add("hidden")
      showError("Failed to upload Excel file. Please check file format.")
    }
  }
})

document.getElementById("generateButton").addEventListener("click", async () => {
  const file = document.getElementById("excelUpload").files[0]
  if (!file) {
    showError("Please upload an Excel file")
    return
  }

  // Set processing state and disable inputs
  isProcessing = true
  toggleInputFields(true)

  document.getElementById("progressSection").classList.remove("hidden")
  document.getElementById("errorMessage").classList.add("hidden")
  const progressBar = document.getElementById("progressBar")
  const progressText = document.getElementById("progressText")
  const resultSection = document.getElementById("resultSection")
  const downloadButton = document.getElementById("downloadButton")
  const cleanupButton = document.getElementById("cleanupButton")
  progressBar.style.width = "0%"
  progressText.textContent = "Starting processing..."

  // FIXED: Clear result section before starting new processing
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
    let lastUpdate = Date.now()
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
            progressBar.style.width = `${data.progress}%`
            progressText.textContent = data.message
            lastUpdate = Date.now()

            if (data.download_ready) {
              resultSection.classList.remove("hidden")
              downloadButton.classList.remove("hidden")
              cleanupButton.classList.remove("hidden")
              downloadButton.onclick = () => {
                window.location.href = `${API_BASE}/download_zip`
              }
            }

            // Load available printers
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

            // Handle print-ready response (replace the existing print-ready handling)
            if (data.print_ready) {
              resultSection.classList.remove("hidden")
              cleanupButton.classList.remove("hidden")

              // Load available printers
              const printers = await loadAvailablePrinters()

              // Create area selection for printing
              const printAreaSelect = document.createElement("select")
              printAreaSelect.id = "printAreaSelect"
              printAreaSelect.className = "p-2 border border-border-medium rounded-lg mr-2"

              // Add options for each area
              printAreaSelect.innerHTML = '<option value="">Select area to print</option>'
              data.areas.forEach((area) => {
                const option = document.createElement("option")
                option.value = area
                option.textContent = area
                printAreaSelect.appendChild(option)
              })

              // Create printer selection dropdown
              const printerSelect = document.createElement("select")
              printerSelect.id = "printerSelect"
              printerSelect.className = "p-2 border border-border-medium rounded-lg mr-2"

              // Add printer options
              printerSelect.innerHTML = '<option value="">Default Printer</option>'
              printers.forEach((printer) => {
                const option = document.createElement("option")
                option.value = printer.name
                option.textContent = printer.name + (printer.is_default ? " (Default)" : "")
                if (printer.is_default) {
                  option.selected = true
                }
                printerSelect.appendChild(option)
              })

              // Create print button
              const printButton = document.createElement("button")
              printButton.id = "printButton"
              printButton.className =
                "btn bg-primary hover:bg-primary-dark text-white py-2 px-4 rounded-lg font-medium flex items-center gap-2"
              printButton.innerHTML = `
    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path>
    </svg>
    Print Selected Area
  `

              // Add event listener for print button
              printButton.addEventListener("click", async () => {
                const selectedArea = printAreaSelect.value
                const selectedPrinter = printerSelect.value

                if (!selectedArea) {
                  showError("Please select an area to print")
                  return
                }

                try {
                  printButton.disabled = true
                  printButton.innerHTML = `
        <div class="loading-spinner-sm"></div>
        Printing...
      `

                  // Build URL with printer parameter if selected
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
                  showSuccess(result.message)
                } catch (error) {
                  showError(`Failed to print: ${error.message}`)
                } finally {
                  printButton.disabled = false
                  printButton.innerHTML = `
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path>
        </svg>
        Print Selected Area
      `
                }
              })

              // Add to result section with improved layout
              const printContainer = document.createElement("div")
              printContainer.className = "mt-4 space-y-3"

              // Create labels and controls
              const areaLabel = document.createElement("label")
              areaLabel.className = "block text-sm font-medium text-text-secondary"
              areaLabel.textContent = "Select Area:"

              const printerLabel = document.createElement("label")
              printerLabel.className = "block text-sm font-medium text-text-secondary"
              printerLabel.textContent = "Select Printer:"

              const controlsContainer = document.createElement("div")
              controlsContainer.className = "flex flex-col sm:flex-row gap-3 items-start sm:items-end"

              const areaContainer = document.createElement("div")
              areaContainer.className = "flex-1"
              areaContainer.appendChild(areaLabel)
              areaContainer.appendChild(printAreaSelect)

              const printerContainer = document.createElement("div")
              printerContainer.className = "flex-1"
              printerContainer.appendChild(printerLabel)
              printerContainer.appendChild(printerSelect)

              const buttonContainer = document.createElement("div")
              buttonContainer.appendChild(printButton)

              controlsContainer.appendChild(areaContainer)
              controlsContainer.appendChild(printerContainer)
              controlsContainer.appendChild(buttonContainer)

              printContainer.appendChild(controlsContainer)
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
    progressText.textContent = "Processing failed."
    progressBar.style.width = "0%"
  } finally {
    // Re-enable inputs when processing is complete
    isProcessing = false
    toggleInputFields(false)
  }
})

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
    tbody.innerHTML = "" // Clear loading state and prior content

    // Update stats
    updateUserStats(users)

    if (users.length === 0) {
      tbody.innerHTML = `
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
      tbody.appendChild(row)
    })

    // Setup search and filter functionality
    setupUserFilters(users)
  } catch (error) {
    tbody = document.getElementById("userTableBody")
    tbody.innerHTML = `
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
    const tbody = document.getElementById("userTableBody")
    tbody.innerHTML = ""

    if (users.length === 0) {
      tbody.innerHTML = `
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
      tbody.appendChild(row)
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
