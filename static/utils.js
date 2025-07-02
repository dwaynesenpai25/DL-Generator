// Utility functions
export function showError(message, isCritical = false) {
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

export function showSuccess(message) {
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

export function toggleInputFields(disabled) {
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

export function getModeColor(mode) {
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

export function updatePageTitle(title, subtitle) {
  document.getElementById("pageTitle").textContent = title
  document.getElementById("pageSubtitle").textContent = subtitle
}

export function showSection(sectionId) {
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
