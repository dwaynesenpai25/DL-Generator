import { CONFIG } from "./config.js"

// API utility functions
export async function apiRequest(endpoint, options = {}) {
  const url = `${CONFIG.API_BASE}${endpoint}`
  const defaultOptions = {
    credentials: "include",
    headers: {
      "Content-Type": "application/json",
      ...options.headers,
    },
  }

  const response = await fetch(url, { ...defaultOptions, ...options })
  return response
}

export async function handleApiResponse(response) {
  if (!response.ok) {
    const errorData = await response.json().catch(() => ({
      detail: `HTTP error ${response.status}`,
    }))
    throw new Error(errorData.detail || `HTTP error ${response.status}`)
  }
  return response.json()
}

// Specific API calls
export async function checkSession() {
  const response = await apiRequest("/check_sessions", { method: "GET" })
  return handleApiResponse(response)
}

export async function logout() {
  const response = await apiRequest("/logout", { method: "GET" })
  return handleApiResponse(response)
}

export async function fetchFolders() {
  const response = await apiRequest("/folders", { method: "GET" })
  return handleApiResponse(response)
}

export async function fetchAllFolders() {
  const response = await apiRequest("/all_folders", { method: "GET" })
  return handleApiResponse(response)
}

export async function fetchDLTypes(folder) {
  const response = await apiRequest("/dl_types", {
    method: "POST",
    body: JSON.stringify({ folder }),
  })
  return handleApiResponse(response)
}

export async function fetchTemplates(folder) {
  const response = await apiRequest("/templates", {
    method: "POST",
    body: JSON.stringify({ folder }),
  })
  return handleApiResponse(response)
}

export async function fetchPlaceholders(folder, dl_type, template) {
  const response = await apiRequest("/placeholders", {
    method: "POST",
    body: JSON.stringify({ folder, dl_type, template }),
  })
  return handleApiResponse(response)
}

export async function fetchTransmittalPlaceholders(folder = null) {
  const endpoint = folder
    ? `/transmittal_placeholders?folder=${encodeURIComponent(folder)}`
    : "/transmittal_placeholders"
  const response = await apiRequest(endpoint, { method: "GET" })
  return handleApiResponse(response)
}

export async function setMode(mode) {
  const response = await apiRequest("/set_mode", {
    method: "POST",
    body: JSON.stringify({ mode }),
  })
  return handleApiResponse(response)
}

export async function setOutputFormat(format) {
  const response = await apiRequest("/set_output_format", {
    method: "POST",
    body: JSON.stringify({ format }),
  })
  return handleApiResponse(response)
}

export async function uploadExcel(formData) {
  const response = await apiRequest("/upload_excel", {
    method: "POST",
    body: formData,
    headers: {}, // Remove Content-Type to let browser set it for FormData
  })
  return handleApiResponse(response)
}

export async function generatePDFs(formData) {
  return apiRequest("/generate_pdfs", {
    method: "POST",
    body: formData,
    headers: {}, // Remove Content-Type to let browser set it for FormData
  })
}

export async function cleanup() {
  const response = await apiRequest("/cleanup", {
    method: "POST",
    body: JSON.stringify({}),
  })
  return handleApiResponse(response)
}

export async function loadAvailablePrinters() {
  try {
    const response = await apiRequest("/printers", { method: "GET" })
    const data = await handleApiResponse(response)
    return data.printers
  } catch (error) {
    console.error("Failed to load printers:", error)
    return []
  }
}

export async function printFiles(area, printer = null) {
  let endpoint = `/print_files/${area}`
  if (printer) {
    endpoint += `?printer=${encodeURIComponent(printer)}`
  }
  const response = await apiRequest(endpoint, { method: "GET" })
  return handleApiResponse(response)
}
