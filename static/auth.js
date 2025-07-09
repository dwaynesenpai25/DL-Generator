import { STATE } from "./config.js"
import { apiRequest, handleApiResponse } from "./api.js"
import { showError } from "./utils.js"
import { updateNavigationAccess, showLoginModal, hideNoClientsModal, showNoClientsModal } from "./ui.js"
import { fetchFolders } from "./document-generator.js"
import watermarkManager from "./watermark-manager.js"

export async function handleLarkCallback(code) {
  if (STATE.isProcessingCallback) {
    console.log("Already processing callback, ignoring duplicate")
    return
  }

  STATE.isProcessingCallback = true
  try {
    const response = await apiRequest(`/lark_callback?code=${code}`, { method: "GET" })
    const data = await handleApiResponse(response)

    console.log("callback", data)
    STATE.currentUser = { username: data.username, role: data.role }
    document.getElementById("loginModal").classList.add("hidden")
    document.getElementById("mainContent").classList.remove("hidden")
    document.getElementById("userDisplay").textContent = `${data.username} (${data.role})`
    window.history.replaceState({}, document.title, "/")

    // Initialize watermark after successful login
    watermarkManager.init(data.username, data.email || data.username)

    await checkSessionStatus()
    updateNavigationAccess()
  } catch (error) {
    console.error("Error in handleLarkCallback:", error)
    document.getElementById("loginError").classList.remove("hidden")
    document.getElementById("loginError").textContent = error.message || "Authentication failed. Please try again."
  } finally {
    STATE.isProcessingCallback = false
  }
}

export async function checkSessionStatus() {
  try {
    const response = await apiRequest("/check_sessions", { method: "GET" })

    if (response.ok) {
      const data = await response.json()
      if (data.success) {
        STATE.currentUser = {
          username: data.username,
          role: data.role,
          access: data.access,
          clients: data.clients || [],
          userInfo: data.avatar.avatar_url || {},
        }

        document.getElementById("loginModal").classList.add("hidden")
        document.getElementById("mainContent").classList.remove("hidden")
        document.getElementById("userDisplay").textContent = `${data.username} (${data.role})`

        // Initialize watermark for existing session
        watermarkManager.init(data.username, data.email || data.username)

        setUserAvatar(STATE.currentUser.userInfo)
        updateNavigationAccess()

        // Check if user has no clients and show modal if needed
        if (!STATE.currentUser.clients || STATE.currentUser.clients.length === 0) {
          showNoClientsModal()
        } else {
          hideNoClientsModal()
          fetchFolders()
        }
      } else {
        showLoginModal()
        // Destroy watermark if session is invalid
        watermarkManager.destroy()
      }
    } else {
      showLoginModal()
      // Destroy watermark if session check fails
      watermarkManager.destroy()
    }
  } catch (error) {
    showLoginModal()
    // Destroy watermark on error
    watermarkManager.destroy()
    console.error("Session check failed:", error)
  }
}

export function setUserAvatar(userInfo) {
  const userAvatar = document.getElementById("userAvatar")
  const userAvatarFallback = document.getElementById("userAvatarFallback")

  if (userInfo) {
    userAvatar.src = userInfo
    userAvatar.classList.remove("hidden")
    userAvatarFallback.classList.add("hidden")
  } else {
    userAvatar.classList.add("hidden")
    userAvatarFallback.classList.remove("hidden")
  }
}

export async function performLogout() {
  try {
    const response = await apiRequest("/logout", { method: "GET" })
    const data = await response.json()

    if (data.success) {
      STATE.currentUser = null

      // Destroy watermark on logout
      watermarkManager.destroy()

      showLoginModal()
      resetUI()
      hideNoClientsModal()
      window.history.replaceState({}, document.title, "/")
    } else {
      showError(data.detail || "Logout failed.")
    }
  } catch (error) {
    showError("Logout failed. Please check if the backend server is running.")
  }
}

export function redirectToLogin() {
  STATE.currentUser = null

  // Destroy watermark on redirect to login
  watermarkManager.destroy()

  showLoginModal()
  resetUI()
  showError("Session expired. Please log in again.")
}

async function resetUI() {
  // This function will be imported from ui.js to avoid circular dependency
  const { resetUI: uiResetUI } = await import("./ui.js")
  uiResetUI()
}
