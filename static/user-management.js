import { apiRequest } from "./api.js"
import { showError, showSuccess } from "./utils.js"
import { redirectToLogin } from "./auth.js"

// Add pagination state at the top of the file
let currentUserPage = 1
const USERS_PER_PAGE = 10

export async function updateUserTable(page = 1) {
    try {
        currentUserPage = page

        // Show loading state
        const tbody = document.getElementById("userTableBody")
        const loadingRow = document.getElementById("loadingState")
        if (loadingRow) {
            loadingRow.classList.remove("hidden")
        }

        const response = await apiRequest("/users", { method: "GET" })

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

        const allUsers = await response.json()
        const userTableBody = document.getElementById("userTableBody")
        userTableBody.innerHTML = "" // Clear loading state and prior content

        // Update stats
        updateUserStats(allUsers)

        // Calculate pagination
        const totalUsers = allUsers.length
        const totalPages = Math.ceil(totalUsers / USERS_PER_PAGE)
        const startIndex = (page - 1) * USERS_PER_PAGE
        const endIndex = startIndex + USERS_PER_PAGE
        const paginatedUsers = allUsers.slice(startIndex, endIndex)

        if (totalUsers === 0) {
            userTableBody.innerHTML = `
        <tr>
          <td colspan="4" class="px-8 py-16 text-center">
            <div class="flex flex-col items-center gap-4">
              <div class="w-16 h-16 bg-slate-100 rounded-2xl flex items-center justify-center">
                <svg class="w-8 h-8 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197m13.5-9a2.5 2.5 0 11-5 0 2.5 2.5 0 015 0z"></path>
                </svg>
              </div>
              <div>
                <h3 class="text-lg font-semibold text-slate-900 mb-1">No users found</h3>
                <p class="text-sm text-slate-500">Get started by adding your first user</p>
              </div>
              <button onclick="document.getElementById('addUserButton').click()" 
                      class="inline-flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white py-2.5 px-4 rounded-lg font-medium transition-colors shadow-sm">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"></path>
                </svg>
                Add First User
              </button>
            </div>
          </td>
        </tr>
      `
            return
        }

        paginatedUsers.forEach((user, index) => {
            const globalIndex = startIndex + index
            const row = document.createElement("tr")
            row.className = "hover:bg-slate-50 transition-all duration-200 group border-b border-slate-100"

            row.innerHTML = `
        <td class="px-6 py-4">
          <div class="flex items-center gap-4">
            <div class="w-12 h-12 bg-gradient-to-br from-green-500 to-green-600 rounded-xl flex items-center justify-center shadow-sm">
              <span class="text-white font-semibold text-sm">
                    ${user.email.substring(0, 2).toUpperCase()}
                    </span>
            </div>
            <div>
              <p class="font-semibold text-slate-900 text-base">${user.email}</p>
              <p class="text-xs text-slate-500">User ID: ${globalIndex + 1}</p>
            </div>
          </div>
        </td>
        <td class="px-6 py-4">
          <div class="flex flex-wrap gap-1.5">
            ${Array.isArray(user.clients)
                    ? user.clients
                        .map(
                            (client) =>
                                `<span class="inline-flex items-center px-2.5 py-1 text-xs font-medium bg-green-50 text-green-700 rounded-full border border-green-200">${client}</span>`,
                        )
                        .join("")
                    : `<span class="inline-flex items-center px-2.5 py-1 text-xs font-medium bg-green-50 text-green-700 rounded-full border border-green-200">${user.clients}</span>`
                }
          </div>
        </td>
        <td class="px-6 py-4">
          <div class="flex items-center gap-3">
            <div class="w-3 h-3 rounded-full ${user.access === "admin" ? "bg-emerald-400" : "bg-green-400"}"></div>
            <span class="inline-flex items-center px-3 py-1.5 text-sm font-medium rounded-full border ${user.access === "admin"
                    ? "bg-emerald-50 text-emerald-700 border-emerald-200"
                    : "bg-slate-50 text-slate-700 border-slate-200"
                }">
              ${user.access === "admin" ? "Administrator" : "Standard User"}
            </span>
          </div>
        </td>
        <td class="px-6 py-4">
          <div class="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity duration-200">
            <button onclick="editUser('${user.email}', decodeURIComponent('${encodeURIComponent(JSON.stringify(user.clients))}'), '${user.access}')"
                class="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-green-600 hover:text-green-800 hover:bg-green-50 rounded-lg transition-all duration-200 border border-transparent hover:border-green-200" 
                title="Edit User">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
              </svg>
              Edit
            </button>
            <button onclick="deleteUser('${user.email}')" 
                class="inline-flex items-center gap-1.5 px-3 py-2 text-sm font-medium text-red-600 hover:text-red-800 hover:bg-red-50 rounded-lg transition-all duration-200 border border-transparent hover:border-red-200" 
                title="Delete User">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5-4h4a1 1 0 011 1v1H9V4a1 1 0 011-1z"></path>
              </svg>
              Delete
            </button>
          </div>
        </td>
      `

            userTableBody.appendChild(row)
        })

        // Render pagination
        renderUserPagination(totalUsers, totalPages, page)

        // Setup search and filter functionality
        setupUserFilters(allUsers)
    } catch (error) {
        const userTableBody = document.getElementById("userTableBody")
        userTableBody.innerHTML = `
      <tr>
        <td colspan="4" class="px-8 py-12 text-center text-red-500">
          <div class="flex flex-col items-center gap-3">
            <svg class="w-12 h-12 text-red-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
            </svg>
            <p class="text-lg font-medium">Failed to load users</p>
            <p class="text-sm">Please try refreshing the page</p>
          </div>
        </td>
      </tr>
    `
        showError("Failed to load users.")
    }
}

// Add new pagination function
function renderUserPagination(totalUsers, totalPages, currentPage) {
    const paginationContainer = document.getElementById("userPagination")
    if (!paginationContainer) return

    let paginationHTML = `
    <div class="flex items-center justify-between bg-white px-6 py-4 border-t border-slate-200">
      <div class="flex items-center gap-2 text-sm text-slate-600">
        <span class="font-medium">Showing</span>
        <span class="px-2 py-1 bg-slate-100 rounded text-slate-900 font-semibold">${Math.min((currentPage - 1) * USERS_PER_PAGE + 1, totalUsers)}</span>
        <span>to</span>
        <span class="px-2 py-1 bg-slate-100 rounded text-slate-900 font-semibold">${Math.min(currentPage * USERS_PER_PAGE, totalUsers)}</span>
        <span>of</span>
        <span class="px-2 py-1 bg-slate-100 rounded text-slate-900 font-semibold">${totalUsers}</span>
        <span>users</span>
      </div>
      <div class="flex items-center gap-1">
  `

    // Previous button
    if (currentPage > 1) {
        paginationHTML += `
      <button onclick="updateUserTable(${currentPage - 1})" 
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
    const startPage = Math.max(1, currentPage - 2)
    const endPage = Math.min(totalPages, currentPage + 2)

    for (let i = startPage; i <= endPage; i++) {
        if (i === currentPage) {
            paginationHTML += `
        <button class="px-4 py-2 text-sm font-semibold text-white bg-green-600 border border-green-600 rounded-lg shadow-sm">
          ${i}
        </button>
      `
        } else {
            paginationHTML += `
        <button onclick="updateUserTable(${i})" 
                class="px-4 py-2 text-sm font-medium text-slate-600 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 hover:text-slate-900 transition-colors">
          ${i}
        </button>
      `
        }
    }

    // Next button
    if (currentPage < totalPages) {
        paginationHTML += `
      <button onclick="updateUserTable(${currentPage + 1})" 
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

export async function deleteUser(email) {
    if (!confirm(`Are you sure you want to delete user ${email}?`)) {
        return
    }

    try {
        const response = await apiRequest(`/users/${encodeURIComponent(email)}`, {
            method: "DELETE",
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
            updateUserTable(currentUserPage) // Refresh the table
            showSuccess("User deleted successfully")
        }
    } catch (error) {
        showError(`Failed to delete user: ${error.message}`)
    }
}

export async function editUser(email, clients, access) {
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

function updateUserStats(users) {
    const totalUsers = users.length
    const adminUsers = users.filter((user) => user.access === "admin").length
    const regularUsers = users.filter((user) => user.access === "user").length

    // Update stat cards
    const totalUsersElement = document.getElementById("totalUsersCount")
    const adminUsersElement = document.getElementById("adminUsersCount")
    const regularUsersElement = document.getElementById("regularUsersCount")

    if (totalUsersElement) totalUsersElement.textContent = totalUsers
    if (adminUsersElement) adminUsersElement.textContent = adminUsers
    if (regularUsersElement) regularUsersElement.textContent = regularUsers
}

function setupUserFilters(allUsers) {
    const searchInput = document.getElementById("userSearchInput")
    const accessFilter = document.getElementById("accessFilter")
    const clearFilters = document.getElementById("clearFilters")

    function filterUsers() {
        const searchTerm = searchInput?.value.toLowerCase() || ""
        const accessLevel = accessFilter?.value || ""

        const filteredUsers = allUsers.filter((user) => {
            const matchesSearch = user.email.toLowerCase().includes(searchTerm)
            const matchesAccess = !accessLevel || user.access === accessLevel
            return matchesSearch && matchesAccess
        })

        renderFilteredUsers(filteredUsers)
    }

    function renderFilteredUsers(users) {
        const userTableBody = document.getElementById("userTableBody")
        userTableBody.innerHTML = ""

        if (users.length === 0) {
            userTableBody.innerHTML = `
        <tr>
          <td colspan="4" class="px-8 py-12 text-center text-gray-500">
            <div class="flex flex-col items-center gap-3">
              <svg class="w-12 h-12 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path>
              </svg>
              <p class="text-lg font-medium">No users match your filters</p>
              <p class="text-sm">Try adjusting your search criteria</p>
            </div>
          </td>
        </tr>
      `
            return
        }

        users.forEach((user, index) => {
            const row = document.createElement("tr")
            row.className = "hover:bg-green-50 transition-all duration-200 ease-in-out group"

            row.innerHTML = `
        <td class="px-8 py-6">
          <div class="flex items-center gap-3">
            <div class="w-10 h-10 bg-gradient-to-br from-green-500 to-green-600 rounded-full flex items-center justify-center">
             <span class="text-white font-semibold text-sm">
                ${user.email.substring(0, 2).toUpperCase()}
                </span>
            </div>
            <div>
              <p class="font-semibold text-gray-900">${user.email}</p>
            </div>
          </div>
        </td>
        <td class="px-8 py-6">
          <div class="flex flex-wrap gap-1">
            ${Array.isArray(user.clients)
                    ? user.clients
                        .map(
                            (client) =>
                                `<span class="inline-flex items-center px-3 py-1 text-xs font-medium bg-green-100 text-green-800 rounded-full border border-green-200">${client}</span>`,
                        )
                        .join("")
                    : `<span class="inline-flex items-center px-3 py-1 text-xs font-medium bg-green-100 text-green-800 rounded-full border border-green-200">${user.clients}</span>`
                }
          </div>
        </td>
        <td class="px-8 py-6">
          <div class="flex items-center gap-2">
            <div class="w-3 h-3 rounded-full ${user.access === "admin" ? "bg-green-400" : "bg-green-400"}"></div>
            <span class="inline-flex items-center px-3 py-1 text-sm font-medium rounded-full ${user.access === "admin"
                    ? "bg-green-100 text-green-800 border border-green-200"
                    : "bg-gray-100 text-gray-800 border border-gray-200"
                }">
              ${user.access === "admin" ? "Administrator" : "Standard User"}
            </span>
          </div>
        </td>
        <td class="px-8 py-6">
          <div class="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity duration-200">
            <button onclick="editUser('${user.email}', decodeURIComponent('${encodeURIComponent(JSON.stringify(user.clients))}'), '${user.access}')"
                class="inline-flex items-center gap-1 px-3 py-2 text-sm font-medium text-green-600 hover:text-green-800 hover:bg-green-50 rounded-lg transition-all duration-200" 
                title="Edit User">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path>
              </svg>
              Edit
            </button>
            <button onclick="deleteUser('${user.email}')" 
                class="inline-flex items-center gap-1 px-3 py-2 text-sm font-medium text-red-600 hover:text-red-800 hover:bg-red-50 rounded-lg transition-all duration-200" 
                title="Delete User">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5-4h4a1 1 0 011 1v1H9V4a1 1 0 011-1z"></path>
              </svg>
              Delete
            </button>
          </div>
        </td>
      `

            userTableBody.appendChild(row)
        })
    }

    // Add event listeners
    if (searchInput) {
        searchInput.addEventListener("input", filterUsers)
    }

    if (accessFilter) {
        accessFilter.addEventListener("change", filterUsers)
    }

    if (clearFilters) {
        clearFilters.addEventListener("click", () => {
            if (searchInput) searchInput.value = ""
            if (accessFilter) accessFilter.value = ""
            renderFilteredUsers(allUsers)
        })
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

// Make functions globally available for onclick handlers
window.deleteUser = deleteUser
window.editUser = editUser

// Make the function globally available
window.updateUserTable = updateUserTable
