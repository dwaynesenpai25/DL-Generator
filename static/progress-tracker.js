export function updateProgressDisplay(data) {
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
