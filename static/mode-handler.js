import { setMode } from "./api.js"
import { showError } from "./utils.js"
import { redirectToLogin } from "./auth.js"
import { clearResultSection } from "./ui.js"

export async function handleModeChange(mode) {
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

    // Reset UI sections
    document.getElementById("selectionSection").classList.add("hidden")
    document.getElementById("transmittalFolderSection").classList.add("hidden")
    document.getElementById("uploadSection").classList.add("hidden")
    document.getElementById("placeholdersDisplay").classList.add("hidden")
    document.getElementById("dataPreview").classList.add("hidden")
    clearResultSection()

    // Keep the mode selection
    document.getElementById("modeSelect").value = mode

    // Handle template status
    if (data.template_status?.transmittal_template) {
      document.getElementById("templateStatusText").textContent = data.template_status.transmittal_template
      document.getElementById("statusDisplay").classList.remove("hidden")
    } else {
      document.getElementById("statusDisplay").classList.add("hidden")
    }

    // Show appropriate sections based on mode
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
