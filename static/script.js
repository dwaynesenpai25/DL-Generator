let currentUser = null
const API_BASE = "http://localhost:5000/api"

// Initialize UI and check session
document.addEventListener("DOMContentLoaded", async () => {
    showSection("dlGeneratorSection")
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

// Check existing session
async function checkSession() {
    try {
        const response = await fetch(`${API_BASE}/check_session`, {
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

    // Try different avatar URLs in order of preference
    //   const avatarUrl = userInfo?.avatar_big || userInfo?.avatar_middle || userInfo?.avatar_thumb || userInfo?.avatar_url

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
async function handleLarkCallback(code) {
    try {
        const response = await fetch(`${API_BASE}/lark_callback?code=${code}`, {
            method: "GET",
            credentials: "include",
        })
        const data = await response.json()
        if (data.success) {
            currentUser = { username: data.username, role: data.role }
            document.getElementById("loginModal").classList.add("hidden")
            document.getElementById("mainContent").classList.remove("hidden")
            document.getElementById("userDisplay").textContent = `${data.username} (${data.role})`
            window.history.replaceState({}, document.title, "/")

            // Check session again to get full user data including access level
            await checkSession()
        } else {
            document.getElementById("loginError").classList.remove("hidden")
            document.getElementById("loginError").textContent = data.detail || "Authentication failed."
        }
    } catch (error) {
        document.getElementById("loginError").classList.remove("hidden")
        document.getElementById("loginError").textContent = "Authentication failed. Please try again."
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

// Reset UI
function resetUI() {
    document.getElementById("modeSelect").value = ""
    document.getElementById("selectionSection").classList.add("hidden")
    document.getElementById("folderSelect").innerHTML = '<option value="">Select Folder</option>'
    document.getElementById("dlTypeSelect").innerHTML = '<option value="">Select DL Type</option>'
    document.getElementById("templateSelect").innerHTML = '<option value="">Select Template</option>'
    document.getElementById("uploadSection").classList.add("hidden")
    document.getElementById("progressSection").classList.add("hidden")
    document.getElementById("errorMessage").classList.add("hidden")
    document.getElementById("placeholdersDisplay").classList.add("hidden")
    document.getElementById("dataPreview").classList.add("hidden")
    document.getElementById("resultSection").classList.add("hidden")
    document.getElementById("progressBar").style.width = "0%"
    document.getElementById("progressText").textContent = ""
    document.getElementById("excelUpload").value = ""
    document.getElementById("dataTable").innerHTML = ""
    document.getElementById("statusDisplay").classList.add("hidden")
    document.getElementById("contentStatus").classList.add("hidden")
    document.getElementById("templatecontentStatus").classList.add("hidden")
}

// Show error
function showError(message) {
    const errorDiv = document.getElementById("errorMessage")
    errorDiv.querySelector("span").textContent = message
    errorDiv.classList.remove("hidden")
    setTimeout(() => errorDiv.classList.add("hidden"), 5000)
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
        if (data.message) {
            document.getElementById("contentStatusText").textContent = data.message
            document.getElementById("contentStatus").classList.remove("hidden")
        }
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
        const placeholdersList = document.getElementById("placeholdersList")
        if (data.message) {
            document.getElementById("templatecontentStatusText").textContent = data.message
            document.getElementById("templatecontentStatus").classList.remove("hidden")
            placeholdersList.innerHTML = ""
            const placeholders = data.placeholders || []
            placeholders.forEach((placeholder) => {
                const li = document.createElement("li")
                li.className = "flex items-center gap-2 text-sm text-text-secondary"
                li.innerHTML = `
          <div class="w-2 h-2 bg-accent rounded-full"></div>
          <code class="bg-background px-2 py-1 rounded text-xs font-mono">${placeholder}</code>
        `
                placeholdersList.appendChild(li)
            })
            document.getElementById("placeholdersDisplay").classList.remove("hidden")
            document.getElementById("uploadSection").classList.remove("hidden")
        } else {
            document.getElementById("templatecontentStatusText").textContent = data.detail
            document.getElementById("templatecontentStatus").classList.remove("hidden")
        }
    } catch (error) {
        showError("Failed to fetch placeholders. Please check template configuration.")
    }
}

// Load audit trail
async function loadAuditTrail() {
    try {
        const response = await fetch(`${API_BASE}/audit_trail`, {
            credentials: "include",
        })
        if (!response.ok) {
            if (response.status === 401) {
                redirectToLogin()
                return
            }
            throw new Error("Failed to fetch audit trail")
        }
        const auditEntries = await response.json()
        const tbody = document.getElementById("auditTableBody")
        tbody.innerHTML = ""

        if (auditEntries.length === 0) {
            tbody.innerHTML =
                '<tr><td colspan="5" class="px-4 py-8 text-center text-text-secondary">No audit entries found</td></tr>'
            return
        }

        auditEntries.forEach((entry) => {
            const row = document.createElement("tr")
            row.className = "hover:bg-surface-hover transition-colors cursor-pointer"
            row.onclick = () => showAuditDetails(entry.id)
            row.innerHTML = `
                <td class="px-4 py-3 text-text-primary">${entry.client}</td>
                <td class="px-4 py-3 text-text-secondary">${entry.processed_by}</td>
                <td class="px-4 py-3 text-text-secondary">${new Date(entry.processed_at).toLocaleString()}</td>
                <td class="px-4 py-3 text-text-secondary">${entry.total_accounts}</td>
                <td class="px-4 py-3">
                    <span class="px-2 py-1 text-xs font-medium bg-primary bg-opacity-10 text-primary rounded-full">
                        ${entry.mode}
                    </span>
                </td>
            `
            tbody.appendChild(row)
        })
    } catch (error) {
        showError("Failed to load audit trail.")
    }
}

// Add new function to show audit details modal
async function showAuditDetails(auditId) {
    try {
        // Show loading indicator
        const loadingModal = document.createElement("div")
        loadingModal.id = "loadingModal"
        loadingModal.className = "fixed inset-0 bg-black bg-opacity-70 flex items-center justify-center z-50 p-4"
        loadingModal.innerHTML = `
      <div class="glass-effect p-6 rounded-2xl shadow-large flex items-center gap-3">
        <div class="animate-spin rounded-full h-8 w-8 border-t-2 border-b-2 border-primary"></div>
        <p class="text-text-primary">Loading account details...</p>
      </div>
    `
        document.body.appendChild(loadingModal)

        // Fetch processed accounts for this audit entry
        const response = await fetch(`${API_BASE}/audit_details/${auditId}`, {
            credentials: "include",
        })

        // Remove loading indicator
        document.body.removeChild(loadingModal)

        if (!response.ok) {
            if (response.status === 401) {
                redirectToLogin()
                return
            }
            throw new Error("Failed to fetch audit details")
        }

        const details = await response.json()

        // Create modal content
        const modalContent = document.createElement("div")
        modalContent.className = "glass-effect p-6 rounded-2xl shadow-large w-full max-w-4xl transform animate-fade-in"
        modalContent.innerHTML = `
      <div class="flex items-center justify-between mb-6">
        <h2 class="text-xl font-semibold text-text-primary">Processed Accounts</h2>
        <button id="closeAuditModal" class="p-2 rounded-lg hover:bg-surface-hover">
          <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path>
          </svg>
        </button>
      </div>
      <div class="mb-4">
        <div class="grid grid-cols-2 gap-4 mb-4">
          <div>
            <p class="text-sm text-text-secondary">Client</p>
            <p class="font-medium">${details.client}</p>
          </div>
          <div>
            <p class="text-sm text-text-secondary">Processed By</p>
            <p class="font-medium">${details.processed_by}</p>
          </div>
          <div>
            <p class="text-sm text-text-secondary">Date</p>
            <p class="font-medium">${new Date(details.processed_at).toLocaleString()}</p>
          </div>
          <div>
            <p class="text-sm text-text-secondary">Mode</p>
            <p class="font-medium">${details.mode}</p>
          </div>
        </div>
      </div>
      <div class="overflow-x-auto max-h-96">
        <table class="min-w-full">
          <thead>
            <tr class="border-b border-border-light">
              <th class="text-left py-3 px-4 font-medium text-text-secondary">DL Code</th>
              <th class="text-left py-3 px-4 font-medium text-text-secondary">Name</th>
              <th class="text-left py-3 px-4 font-medium text-text-secondary">Address</th>
              <th class="text-left py-3 px-4 font-medium text-text-secondary">Area</th>
            </tr>
          </thead>
          <tbody id="accountDetailsList" class="divide-y divide-border-light">
            ${details.accounts.length > 0
                ? details.accounts
                    .map(
                        (account) => `
                <tr class="hover:bg-surface-hover transition-colors">
                  <td class="px-4 py-3 text-text-primary">${account.dl_code || ""}</td>
                  <td class="px-4 py-3 text-text-secondary">${account.name || ""}</td>
                  <td class="px-4 py-3 text-text-secondary">${account.address || ""}</td>
                  <td class="px-4 py-3 text-text-secondary">${account.area || ""}</td>
                </tr>
              `,
                    )
                    .join("")
                : '<tr><td colspan="4" class="px-4 py-8 text-center text-text-secondary">No account details available</td></tr>'
            }
          </tbody>
        </table>
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
            document.body.removeChild(modalContainer)
        })

        // Close on outside click
        modalContainer.addEventListener("click", (e) => {
            if (e.target === modalContainer) {
                document.body.removeChild(modalContainer)
            }
        })
    } catch (error) {
        console.error("Failed to load audit details:", error)
        showError("Failed to load audit details.")
    }
}

// Refresh audit trail
document.getElementById("refreshAuditButton").addEventListener("click", () => {
    loadAuditTrail()
})

// Event listeners for DL Generator
document.getElementById("modeSelect").addEventListener("change", async (e) => {
    const mode = e.target.value
    if (!mode) {
        resetUI()
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
        resetUI()
        document.getElementById("modeSelect").value = mode
        if (data.template_status?.transmittal_template) {
            document.getElementById("templateStatusText").textContent = data.template_status.transmittal_template
            document.getElementById("statusDisplay").classList.remove("hidden")
        } else {
            document.getElementById("statusDisplay").classList.add("hidden")
        }
        if (mode === "Transmittal Only") {
            document.getElementById("uploadSection").classList.remove("hidden")
        } else {
            document.getElementById("selectionSection").classList.remove("hidden")
            await fetchFolders()
        }
    } catch (error) {
        showError("Failed to set mode. Please check if the backend server is running.")
        document.getElementById("modeSelect").value = ""
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

document.getElementById("templateSelect").addEventListener("change", (e) => {
    if (e.target.value) {
        fetchPlaceholders(
            document.getElementById("folderSelect").value,
            document.getElementById("dlTypeSelect").value,
            e.target.value,
        )
    }
})

document.getElementById("excelUpload").addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (file) {
        const formData = new FormData()
        formData.append("file", file)
        try {
            const response = await fetch(`${API_BASE}/upload_excel`, {
                method: "POST",
                body: formData,
            })
            const data = await response.json()
            if (data.error) {
                showError(data.error)
                return
            }

            const tableContainer = document.getElementById("dataTable");
            tableContainer.innerHTML = "";
            tableContainer.className = "max-h-[500px] overflow-auto rounded-lg border border-border-light";

            // Create table
            const table = document.createElement("table");
            table.className = "min-w-full text-sm";

            // Create thead with sticky header
            const thead = document.createElement("thead");
            const headerRow = document.createElement("tr");
            headerRow.className = "border-b border-border-light bg-white";

            Object.keys(data.data[0]).forEach((key) => {
                const th = document.createElement("th");
                th.className = "sticky top-0 bg-white z-10 text-left py-3 px-4 font-medium text-text-secondary";
                th.textContent = key;
                headerRow.appendChild(th);
            });
            thead.appendChild(headerRow);
            table.appendChild(thead);

            // Create tbody
            const tbody = document.createElement("tbody");
            tbody.className = "divide-y divide-border-light";

            data.data.forEach((row) => {
                const tr = document.createElement("tr");
                tr.className = "hover:bg-surface-hover transition-colors";

                Object.values(row).forEach((value) => {
                    const td = document.createElement("td");
                    td.className = "px-4 py-3 text-text-secondary";
                    td.textContent = value;
                    tr.appendChild(td);
                });

                tbody.appendChild(tr);
            });

            table.appendChild(tbody);
            tableContainer.appendChild(table);

            // Show data preview section
            document.getElementById("dataPreview").classList.remove("hidden");

        } catch (error) {
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
    document.getElementById("progressSection").classList.remove("hidden")
    document.getElementById("errorMessage").classList.add("hidden")
    const progressBar = document.getElementById("progressBar")
    const progressText = document.getElementById("progressText")
    const resultSection = document.getElementById("resultSection")
    const downloadButton = document.getElementById("downloadButton")
    const cleanupButton = document.getElementById("cleanupButton")
    progressBar.style.width = "0%"
    progressText.textContent = "Starting processing..."
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
        const tbody = document.getElementById("userTableBody")
        tbody.innerHTML = ""

        if (users.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" class="px-4 py-8 text-center text-text-secondary">No users found</td></tr>'
            return
        }

        users.forEach((user) => {
            const row = document.createElement("tr")
            row.className = "hover:bg-surface-hover transition-colors"
            const clientsDisplay = Array.isArray(user.clients) ? user.clients.join(", ") : user.clients
            row.innerHTML = `
                <td class="px-4 py-3 text-text-primary">${user.email}</td>
               <td class="px-4 py-3 text-text-secondary">
                ${Array.isArray(user.clients)
                    ? user.clients.map(client => `<span class="inline-block bg-blue-100 text-blue-800 text-xs font-semibold mr-1 px-2.5 py-0.5 rounded">${client}</span>`).join('')
                    : `<span class="inline-block bg-blue-100 text-blue-800 text-xs font-semibold px-2.5 py-0.5 rounded">${user.clients}</span>`
                }
                </td>
                <td class="px-4 py-3">
                    <span class="px-2 py-1 text-xs font-medium rounded-full ${user.access === "admin" ? "bg-blue-100 text-blue-800" : "bg-gray-100 text-gray-800"
                }">
                        ${user.access}
                    </span>
                </td>
                <td class="px-4 py-3">
                    <div class="flex items-center gap-2">
                       <button onclick="editUser('${user.email}', decodeURIComponent('${encodeURIComponent(JSON.stringify(user.clients))}'), '${user.access}')"

                            class="icon-btn p-2 text-primary hover:bg-primary hover:bg-opacity-10 rounded-lg transition-colors" 
                            title="Edit User">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
                            </svg>
                        </button>
                        <button onclick="deleteUser('${user.email}')" 
                            class="icon-btn p-2 text-error hover:bg-error hover:bg-opacity-10 rounded-lg transition-colors" 
                            title="Delete User">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5-4h4a1 1 0 011 1v1H9V4a1 1 0 011-1z"></path>
                            </svg>
                        </button>
                    </div>
                </td>
            `
            tbody.appendChild(row)
        })
    } catch (error) {
        showError("Failed to load users.")
    }
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

// Enable drag and drop for Excel file
const dropZone = document.querySelector("label[for='excelUpload']")
if (dropZone) {
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault()
        dropZone.classList.add("border-primary", "bg-primary", "bg-opacity-5")
    })
    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("border-primary", "bg-primary", "bg-opacity-5")
    })
    dropZone.addEventListener("drop", (e) => {
        e.preventDefault()
        dropZone.classList.remove("border-primary", "bg-primary", "bg-opacity-5")
        document.getElementById("excelUpload").files = e.dataTransfer.files
        document.getElementById("excelUpload").dispatchEvent(new Event("change"))
    })
}

// Make functions globally available
window.deleteUser = deleteUser
window.editUser = editUser
