import { printFiles } from "./api.js"

export function createPrintControls(areas, printers) {
  const printContainer = document.createElement("div")
  printContainer.className = "mt-6 bg-gradient-to-r from-slate-50 to-gray-100 rounded-lg border-l-4 border-blue-500"
  printContainer.innerHTML = `
    <div class="p-5">
      <div class="flex items-center justify-between mb-5">
        <div class="flex items-center gap-2">
          <span class="text-2xl">🖨️</span>
          <h4 class="text-lg font-bold text-gray-900">Print Documents</h4>
        </div>
        <div class="h-2 w-2 bg-blue-500 rounded-full animate-pulse"></div>
      </div>
      
      <div class="bg-white rounded-lg p-4 shadow-inner border border-gray-200">
        <div class="flex flex-col md:flex-row gap-3">
          <div class="flex-1 min-w-0">
            <div class="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1">Area</div>
            <select id="printAreaSelect" class="w-full p-2 text-sm border-2 border-gray-200 rounded-md focus:border-blue-400 focus:outline-none">
              <option value="">Select area...</option>
              ${areas.map((area) => `<option value="${area}">${area}</option>`).join("")}
            </select>
          </div>
          
          <div class="flex-1 min-w-0">
            <div class="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1">Printer</div>
            <select id="printerSelect" class="w-full p-2 text-sm border-2 border-gray-200 rounded-md focus:border-blue-400 focus:outline-none">
              <option value="">Default</option>
              ${printers
                .map(
                  (printer) => `
                <option value="${printer.name}" ${printer.is_default ? "selected" : ""}>
                  ${printer.name}${printer.is_default ? " ⭐" : ""}
                </option>
              `,
                )
                .join("")}
            </select>
          </div>
          
          <div class="md:w-auto">
            <div class="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-1">&nbsp;</div>
            <button id="printButton" class="w-full md:w-auto px-6 py-2 bg-blue-500 hover:bg-blue-600 text-white text-sm font-medium rounded-md transition-colors duration-150 flex items-center gap-2">
              <span>▶</span>
              Print
            </button>
          </div>
        </div>
      </div>
      
      <div id="printStatus" class="hidden mt-3"></div>
    </div>
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
      showPrintStatus("⚠️ Select an area first", "warning")
      return
    }

    try {
      printButton.disabled = true
      printButton.innerHTML = `<span class="animate-spin">⟳</span> Printing...`

      const result = await printFiles(selectedArea, selectedPrinter)
      showPrintStatus(`✓ ${result.message}`, "success")
    } catch (error) {
      showPrintStatus(`✗ ${error.message}`, "error")
    } finally {
      printButton.disabled = false
      printButton.innerHTML = `<span>▶</span> Print`
    }
  })

  function showPrintStatus(message, type) {
    const colors = {
      success: "bg-green-100 text-green-800 border-green-300",
      error: "bg-red-100 text-red-800 border-red-300",
      warning: "bg-yellow-100 text-yellow-800 border-yellow-300"
    }
    
    printStatus.className = `mt-3 p-2 text-sm rounded border ${colors[type]} animate-fade-in`
    printStatus.textContent = message
    printStatus.classList.remove("hidden")
    
    setTimeout(() => {
      printStatus.classList.add("hidden")
    }, 4000)
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