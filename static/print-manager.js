import { printFiles } from "./api.js"

export function createPrintControls(areas, printers) {
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

      const result = await printFiles(selectedArea, selectedPrinter)
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

export async function loadAvailablePrinters() {
  try {
    const { apiRequest, handleApiResponse } = await import("./api.js")
    const response = await apiRequest("/printers", { method: "GET" })
    const data = await handleApiResponse(response)
    return data.printers
  } catch (error) {
    console.error("Failed to load printers:", error)
    return []
  }
}
