import { CONFIG, STATE } from "./config.js"
import { apiRequest } from "./api.js"
import { showError, getModeColor } from "./utils.js"
import { redirectToLogin } from "./auth.js"

export async function loadAuditTrail(page = 1) {
  try {
    STATE.currentAuditPage = page
    const response = await apiRequest(`/audit_trail?page=${page}&limit=${CONFIG.PAGINATION_LIMIT}`, {
      method: "GET",
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
      tbody.innerHTML = `
        <tr>
          <td colspan="5" class="px-8 py-16 text-center">
            <div class="flex flex-col items-center gap-4">
              <div class="w-16 h-16 bg-slate-100 rounded-2xl flex items-center justify-center">
                <svg class="w-8 h-8 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v10a2 2 0 002 2h8a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path>
                </svg>
              </div>
              <div>
                <h3 class="text-lg font-semibold text-slate-900 mb-1">No audit entries found</h3>
                <p class="text-sm text-slate-500">Start generating documents to see audit trail</p>
              </div>
            </div>
          </td>
        </tr>
      `
      return
    }

    data.entries.forEach((entry, index) => {
      const row = document.createElement("tr")
      row.className = "hover:bg-slate-50 transition-all duration-200 cursor-pointer border-b border-slate-100 group"
      row.onclick = () => showAuditDetails(entry.id)

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
          <div class="flex items-center gap-4">
            <div class="w-12 h-12 bg-gradient-to-br from-green-500 to-green-600 rounded-xl flex items-center justify-center shadow-sm">
              <svg class="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"></path>
              </svg>
            </div>
            <div>
              <p class="font-semibold text-slate-900 text-base">${entry.client}</p>
              <p class="text-xs text-slate-500 flex items-center gap-1 mt-1">
                <svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"></path>
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"></path>
                </svg>
                Click to view details
              </p>
            </div>
          </div>
        </td>
        <td class="px-6 py-4">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-slate-100 rounded-full flex items-center justify-center">
              <svg class="w-5 h-5 text-slate-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"></path>
              </svg>
            </div>
            <div>
              <p class="font-medium text-slate-900">${entry.processed_by}</p>
              <p class="text-xs text-slate-500">Administrator</p>
            </div>
          </div>
        </td>
        <td class="px-6 py-4">
          <div class="text-slate-700">
            <p class="font-medium text-sm">${formattedDate}</p>
            <p class="text-xs text-slate-500">${formattedTime}</p>
          </div>
        </td>
        <td class="px-6 py-4">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-emerald-100 rounded-full flex items-center justify-center">
              <svg class="w-5 h-5 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path>
              </svg>
            </div>
            <div>
              <p class="font-bold text-lg text-slate-900">${entry.total_accounts}</p>
              <p class="text-xs text-slate-500">accounts</p>
            </div>
          </div>
        </td>
        <td class="px-6 py-4">
          <span class="inline-flex items-center px-3 py-1.5 text-xs font-semibold rounded-full border ${getModeColor(entry.mode)}">
            ${entry.mode}
          </span>
        </td>
      `

      tbody.appendChild(row)
    })

    // Render enhanced pagination
    renderAuditPagination(data.pagination)
  } catch (error) {
    showError("Failed to load audit trail.")
  }
}

function renderAuditPagination(pagination) {
  const paginationContainer = document.getElementById("auditPagination")
  if (!paginationContainer) return

  let paginationHTML = `
    <div class="flex items-center justify-between bg-white px-6 py-4 border-t border-slate-200">
      <div class="flex items-center gap-2 text-sm text-slate-600">
        <span class="font-medium">Showing</span>
        <span class="px-2 py-1 bg-slate-100 rounded text-slate-900 font-semibold">${pagination.current_page}</span>
        <span>of</span>
        <span class="px-2 py-1 bg-slate-100 rounded text-slate-900 font-semibold">${pagination.total_pages}</span>
        <span>pages (${pagination.total_count} total entries)</span>
      </div>
      <div class="flex items-center gap-1">
  `

  // Previous button
  if (pagination.has_prev) {
    paginationHTML += `
      <button onclick="loadAuditTrail(${pagination.current_page - 1})" 
              class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 hover:text-slate-900 transition-colors">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7"></path>
        </svg>
        Previous
      </button>
    `
  } else {
    paginationHTML += `
      <button disabled class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-400 bg-slate-100 border border-slate-200 rounded-lg cursor-not-allowed">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 19l-7-7 7-7"></path>
        </svg>
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
        <button class="px-4 py-2 text-sm font-semibold text-white bg-green-600 border border-green-600 rounded-lg shadow-sm">
          ${i}
        </button>
      `
    } else {
      paginationHTML += `
        <button onclick="loadAuditTrail(${i})" 
                class="px-4 py-2 text-sm font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 hover:text-slate-900 transition-colors">
          ${i}
        </button>
      `
    }
  }

  // Next button
  if (pagination.has_next) {
    paginationHTML += `
      <button onclick="loadAuditTrail(${pagination.current_page + 1})" 
              class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 hover:text-slate-900 transition-colors">
        Next
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path>
        </svg>
      </button>
    `
  } else {
    paginationHTML += `
      <button disabled class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-400 bg-slate-100 border border-slate-200 rounded-lg cursor-not-allowed">
        Next
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path>
        </svg>
      </button>
    `
  }

  paginationHTML += `
      </div>
    </div>
  `

  paginationContainer.innerHTML = paginationHTML
}

export async function showAuditDetails(auditId, page = 1) {
  try {
    STATE.currentAuditDetailsPage = page

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
    const response = await apiRequest(`/audit_details/${auditId}?page=${page}&limit=${CONFIG.BATCH_SIZE}`, {
      method: "GET",
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
        <div class="bg-green-50 p-4 rounded-xl border border-green-200">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-green-100 rounded-lg flex items-center justify-center">
              <svg class="w-5 h-5 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"></path>
              </svg>
            </div>
            <div>
              <p class="text-sm text-green-600 font-medium">Client</p>
              <p class="font-bold text-green-900">${details.client}</p>
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
        <div class="bg-green-50 p-4 rounded-xl border border-green-200">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-green-100 rounded-lg flex items-center justify-center">
              <svg class="w-5 h-5 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path>
              </svg>
            </div>
            <div>
              <p class="text-sm text-green-600 font-medium">Date</p>
              <p class="font-bold text-green-900">${new Date(details.processed_at).toLocaleDateString()}</p>
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

// Make functions globally available for onclick handlers
window.loadAuditTrail = loadAuditTrail
window.showAuditDetails = showAuditDetails
