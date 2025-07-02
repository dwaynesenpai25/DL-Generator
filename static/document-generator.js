import { CONFIG, STATE } from "./config.js"
import {
  fetchFolders as apiFetchFolders,
  fetchDLTypes,
  fetchTemplates,
  fetchPlaceholders,
  uploadExcel,
  generatePDFs,
  cleanup,
} from "./api.js"
import { showError, toggleInputFields } from "./utils.js"
import { redirectToLogin } from "./auth.js"
import { showTemplateCombinedAlert, clearResultSection } from "./ui.js"
import { updateProgressDisplay } from "./progress-tracker.js"
import { loadAvailablePrinters, createPrintControls } from "./print-manager.js"

export async function fetchFolders() {
  try {
    const folders = await apiFetchFolders()
    const folderSelect = document.getElementById("folderSelect")
    folderSelect.innerHTML = '<option value="">Select Folder</option>'

    folders.forEach((folder) => {
      const option = document.createElement("option")
      option.value = folder
      option.textContent = folder
      folderSelect.appendChild(option)
    })

    if (STATE.currentUser && STATE.currentUser.access === "admin1") {
      await loadTemplateFoldersForModal()
    }
  } catch (error) {
    console.error("fetchFolders error:", error)
    if (error.message.includes("401")) {
      redirectToLogin()
      return
    }
    showError(`Failed to fetch folders: ${error.message}. Please check FTP configuration or server status.`)
  }
}

export async function handleDLTypeSelection(folder) {
  try {
    const dlTypes = await fetchDLTypes(folder)
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
    if (error.message.includes("401")) {
      redirectToLogin()
      return
    }
    if (error.message.includes("403")) {
      showError("Access denied to this template folder.")
      return
    }
    showError(`Failed to fetch DL types for "${folder}": ${error.message}.`)
  }
}

export async function handleTemplateSelection(folder) {
  try {
    const data = await fetchTemplates(folder)
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
    if (error.message.includes("401")) {
      redirectToLogin()
      return
    }
    if (error.message.includes("403")) {
      showError("Access denied to this template folder.")
      return
    }
    showError(`Failed to fetch templates: ${error.message}. Please check FTP connection or server status.`)
  }
}

export async function handlePlaceholderFetch(folder, dl_type, template) {
  document.getElementById("placeholdersLoadingOverlay").classList.remove("hidden")

  try {
    const data = await fetchPlaceholders(folder, dl_type, template)
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
    if (error.message.includes("401")) {
      redirectToLogin()
      return
    }
    if (error.message.includes("403")) {
      showError("Access denied to this template folder.")
      return
    }
    showError("Failed to fetch placeholders. Please check template configuration.")
  } finally {
    document.getElementById("placeholdersLoadingOverlay").classList.add("hidden")
  }
}

export async function handleExcelUpload(file) {
  document.getElementById("excelLoadingOverlay").classList.remove("hidden")
  const formData = new FormData()
  formData.append("file", file)

  try {
    const data = await uploadExcel(formData)
    document.getElementById("excelLoadingOverlay").classList.add("hidden")

    if (data && data.data && Array.isArray(data.data) && data.data.length > 0) {
      const totalRows = data.data.length
      const previewRows = data.data.slice(0, CONFIG.PREVIEW_ROWS_LIMIT)

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
              ${totalRows > CONFIG.PREVIEW_ROWS_LIMIT ? `(showing first ${CONFIG.PREVIEW_ROWS_LIMIT} rows in preview)` : ""}
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

      // Add "more rows" indicator if there are more than preview limit
      if (totalRows > CONFIG.PREVIEW_ROWS_LIMIT) {
        const moreRowsIndicator = document.createElement("tr")
        moreRowsIndicator.className = "bg-gray-100"
        const td = document.createElement("td")
        td.colSpan = Object.keys(data.data[0]).length
        td.className = "px-4 py-3 text-center text-gray-500 italic"
        td.textContent = `... and ${totalRows - CONFIG.PREVIEW_ROWS_LIMIT} more rows`
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

export async function handleDocumentGeneration(file) {
  // Set processing state and prevent reloads
  STATE.isProcessing = true
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
      reject(new Error(`Processing timed out after ${CONFIG.TIMEOUT_DURATION / 1000} seconds`))
    }, CONFIG.TIMEOUT_DURATION)
  })

  try {
    const response = await Promise.race([generatePDFs(formData), timeoutPromise])

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
                window.location.href = `${CONFIG.API_BASE}/download_zip`
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
    STATE.isProcessing = false
    toggleInputFields(false)
  }
}

export async function handleCleanup() {
  try {
    const data = await cleanup()
    if (data.success) {
      document.getElementById("resultSection").classList.add("hidden")
      document.getElementById("progressSection").classList.add("hidden")
      const { resetUI } = await import("./ui.js")
      resetUI()
    } else {
      showError(data.detail || "Failed to cleanup files")
    }
  } catch (error) {
    if (error.message.includes("401")) {
      redirectToLogin()
      return
    }
    showError("Failed to cleanup files. Please check backend server.")
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
