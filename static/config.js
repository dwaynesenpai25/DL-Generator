// Configuration and constants
export const CONFIG = {
  API_BASE: "http://172.20.0.86:5000/api",
  // API_BASE: "http://localhost:5000/api", // Alternative endpoint
  TIMEOUT_DURATION: 120000, // 2 minutes
  PAGINATION_LIMIT: 10,
  PREVIEW_ROWS_LIMIT: 10,
  BATCH_SIZE: 50,
}

// Global state
export const STATE = {
  currentUser: null,
  isProcessing: false,
  isProcessingCallback: false,
  currentAuditPage: 1,
  currentAuditDetailsPage: 1,
}
