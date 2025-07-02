import { fetchTransmittalPlaceholders } from "./api.js"
import { showError } from "./utils.js"
import { redirectToLogin } from "./auth.js"

export async function handleTransmittalPlaceholders() {
  document.getElementById("placeholdersLoadingOverlay").classList.remove("hidden")

  try {
    const data = await fetchTransmittalPlaceholders()
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
    if (error.message.includes("401")) {
      redirectToLogin()
      return
    }
    showError("Failed to fetch transmittal placeholders. Please check template configuration.")
  } finally {
    document.getElementById("placeholdersLoadingOverlay").classList.add("hidden")
  }
}
