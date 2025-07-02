import { CONFIG } from "./config.js"
import {
  handleDLTypeSelection,
  handleTemplateSelection,
  handlePlaceholderFetch,
  handleExcelUpload,
  handleDocumentGeneration,
  handleCleanup,
} from "./document-generator.js"
import { setMode, setOutputFormat, fetchTransmittalPlaceholders } from "./api.js"
import { showError } from "./utils.js"
import { redirectToLogin } from "./auth.js"
import { performLogout } from "./auth.js"

export function setupEventHandlers() {
  // Authentication handlers
  document.getElementById("loginButton").addEventListener("click", () => {
    window.location.href = `${CONFIG.API_BASE}/login`
  })

  document.getElementById("logoutButton").addEventListener("click", performLogout)

  // Document generation handlers
  document.getElementById("outputFormatSelect").addEventListener("change", async (e) => {
    const format = e.target.value
    if (!format) {
      document.querySelector(".card:has(#modeSelect)").classList.add("hidden")
      return
    }

    try {
      await setOutputFormat(format)
      document.querySelector(".card:has(#modeSelect)").classList.remove("hidden")

      if (format === "print") {
        document.getElementById("printFormatInfo").classList.remove("hidden")
        document.getElementById("zipFormatInfo").classList.add("hidden")
      } else {
        document.getElementById("printFormatInfo").classList.add("hidden")
        document.getElementById("zipFormatInfo").classList.remove("hidden")
      }
    } catch (error) {
      if (error.message.includes("401")) {
        redirectToLogin()
        return
      }
      showError("Failed to set output format. Please try again.")
    }
  })

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
      const data = await setMode(mode)
      document.getElementById("selectionSection").classList.add("hidden")
      document.getElementById("transmittalFolderSection").classList.add("hidden")
      document.getElementById("uploadSection").classList.add("hidden")
      document.getElementById("placeholdersDisplay").classList.add("hidden")
      document.getElementById("dataPreview").classList.add("hidden")

      const { clearResultSection } = await import("./ui.js")
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
        const { fetchFolders } = await import("./document-generator.js")
        await fetchFolders()
      } else {
        document.getElementById("selectionSection").classList.remove("hidden")
        const { fetchFolders } = await import("./document-generator.js")
        await fetchFolders()
      }
    } catch (error) {
      if (error.message.includes("401")) {
        redirectToLogin()
        return
      }
      showError("Failed to set mode. Please check if the backend server is running.")
      document.getElementById("modeSelect").value = ""
    }
  })

  document.getElementById("folderSelect").addEventListener("change", (e) => {
    if (e.target.value) {
      handleDLTypeSelection(e.target.value)
    }
  })

  document.getElementById("dlTypeSelect").addEventListener("change", (e) => {
    if (e.target.value) {
      handleTemplateSelection(document.getElementById("folderSelect").value)
    }
  })

  document.getElementById("templateSelect").addEventListener("change", async (e) => {
    if (e.target.value) {
      const mode = document.getElementById("modeSelect").value
      const folder = document.getElementById("folderSelect").value
      const dlType = document.getElementById("dlTypeSelect").value
      const template = e.target.value

      await handlePlaceholderFetch(folder, dlType, template)

      if (mode === "DL w/ Transmittal") {
        document.getElementById("placeholdersLoadingOverlay").classList.remove("hidden")
        try {
          const data = await fetchTransmittalPlaceholders()
          const placeholdersList = document.getElementById("placeholdersList")
          const separator = document.createElement("li")
          separator.className = "py-2 border-t border-border-light mt-2 pt-2"
          separator.innerHTML = `
            <div class="flex items-center gap-2">
              <div class="w-4 h-4 bg-green-500 rounded-full"></div>
              <span class="font-medium text-green-700">Transmittal Placeholders</span>
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
              li.className = "flex items-center gap-2 text-sm text-green-600"
              li.innerHTML = `
                <div class="w-2 h-2 bg-green-500 rounded-full"></div>
                <code class="bg-green-50 px-2 py-1 rounded text-xs font-mono">${placeholder}</code>
              `
              placeholdersList.appendChild(li)
            })
          } else {
            const li = document.createElement("li")
            li.className = "text-sm text-green-600"
            li.textContent = "No transmittal placeholders found."
            placeholdersList.appendChild(li)
          }
        } catch (error) {
          console.error("Error fetching transmittal placeholders:", error)
        } finally {
          document.getElementById("placeholdersLoadingOverlay").classList.add("hidden")
        }
      }
    }
  })

  document.getElementById("transmittalFolderSelect").addEventListener("change", async (e) => {
    const folder = e.target.value
    if (folder) {
      try {
        const data = await fetchTransmittalPlaceholders(folder)
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
        if (error.message.includes("401")) {
          redirectToLogin()
          return
        }
        showError("Failed to fetch transmittal placeholders. Please check template configuration.")
      }
    } else {
      document.getElementById("placeholdersDisplay").classList.add("hidden")
      document.getElementById("uploadSection").classList.add("hidden")
    }
  })

  document.getElementById("excelUpload").addEventListener("change", async (e) => {
    const file = e.target.files[0]
    if (file) {
      await handleExcelUpload(file)
    }
  })

  document.getElementById("generateButton").addEventListener("click", async () => {
    const file = document.getElementById("excelUpload").files[0]
    if (!file) {
      showError("Please upload an Excel file")
      return
    }
    await handleDocumentGeneration(file)
  })

  document.getElementById("cleanupButton").addEventListener("click", handleCleanup)

  // User management handlers
  document.getElementById("addUserButton")?.addEventListener("click", async () => {
    document.getElementById("userModalTitle").textContent = "Add New User"
    document.getElementById("modalEmail").value = ""
    document.getElementById("modalAccess").value = "user"
    document.getElementById("modalEmail").disabled = false

    // Clear the original email dataset
    delete document.getElementById("userForm").dataset.originalEmail

    // Always load template folders for the modal
    const { updateUserTable } = await import("./user-management.js")
    await loadTemplateFoldersForModal()

    // Clear all checkboxes
    const checkboxes = document.querySelectorAll('#modalClients input[type="checkbox"]')
    checkboxes.forEach((checkbox) => {
      checkbox.checked = false
    })

    document.getElementById("userModal").classList.remove("hidden")
  })

  document.getElementById("cancelUserModal")?.addEventListener("click", () => {
    document.getElementById("userModal").classList.add("hidden")
  })

  document.getElementById("userForm")?.addEventListener("submit", async (e) => {
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
      const { apiRequest } = await import("./api.js")
      let response

      if (originalEmail) {
        // Update existing user
        response = await apiRequest(`/users/${encodeURIComponent(originalEmail)}`, {
          method: "PUT",
          body: JSON.stringify({
            email,
            clients: selectedClients,
            access,
          }),
        })
      } else {
        // Create new user
        response = await apiRequest("/users", {
          method: "POST",
          body: JSON.stringify({
            email,
            clients: selectedClients,
            access,
          }),
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

        const { updateUserTable } = await import("./user-management.js")
        updateUserTable() // Refresh the table

        const { showSuccess } = await import("./utils.js")
        showSuccess(`User ${originalEmail ? "updated" : "created"} successfully`)
      }
    } catch (error) {
      showError(`Failed to ${originalEmail ? "update" : "create"} user: ${error.message}`)
    }
  })

  // Refresh audit trail
  document.getElementById("refreshAuditButton")?.addEventListener("click", () => {
    import("./audit-trail.js").then(({ loadAuditTrail }) => {
      loadAuditTrail(1) // Reset to first page
    })
  })

  // Enable drag and drop for Excel file
  setupDragAndDrop()
}

async function fetchFoldersForTransmittal() {
  try {
    const { fetchFolders } = await import("./api.js")
    const folders = await fetchFolders()
    const transmittalFolderSelect = document.getElementById("transmittalFolderSelect")
    transmittalFolderSelect.innerHTML = '<option value="">Select Client Folder</option>'

    folders.forEach((folder) => {
      const option = document.createElement("option")
      option.value = folder
      option.textContent = folder
      transmittalFolderSelect.appendChild(option)
    })
  } catch (error) {
    if (error.message.includes("401")) {
      redirectToLogin()
      return
    }
    showError("Failed to fetch folders. Please check FTP configuration.")
  }
}

async function loadTemplateFoldersForModal() {
  try {
    const { fetchAllFolders } = await import("./api.js")
    const folders = await fetchAllFolders()
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
  } catch (error) {
    console.error("Failed to load template folders for modal:", error)
  }
}

function setupDragAndDrop() {
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

      import("./config.js").then(({ STATE }) => {
        if (!STATE.isProcessing) {
          document.getElementById("excelUpload").files = e.dataTransfer.files
          document.getElementById("excelUpload").dispatchEvent(new Event("change"))
        }
      })
    })
  }
}
