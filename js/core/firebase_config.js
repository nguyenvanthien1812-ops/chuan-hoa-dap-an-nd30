/**
 * firebase_config.js - Cấu hình Firebase & Danh sách Admin cho DocNormalizer Pro
 */

(function(window) {
    // 1. Danh sách email Quản trị viên (Admin) cao nhất
    // Các tài khoản này khi đăng nhập Google sẽ TỰ ĐỘNG có quyền Admin và bản quyền Vĩnh viễn.
    const DEFAULT_ADMIN_EMAILS = [
        "nguyenvanthien1812@gmail.com"
    ];

    // 2. Cấu hình Firebase mặc định (Bạn có thể dán thông tin từ Firebase Console vào đây hoặc nhập qua giao diện Web)
    const DEFAULT_FIREBASE_CONFIG = {
        apiKey: "AIzaSyDlI5L2wx0s-ZqBs8coxeXpJmSctZdR8l8",
        authDomain: "chuan-hoa-dap-an-nd30.firebaseapp.com",
        projectId: "chuan-hoa-dap-an-nd30",
        storageBucket: "chuan-hoa-dap-an-nd30.firebasestorage.app",
        messagingSenderId: "335548189735",
        appId: "1:335548189735:web:8dbf61064c4ea7f7cf1202",
        measurementId: "G-NCK8PWJLW8"
    };

    const STORAGE_KEY = "docnormalizer_firebase_config";

    function getFirebaseConfig() {
        try {
            const stored = localStorage.getItem(STORAGE_KEY);
            if (stored) {
                const parsed = JSON.parse(stored);
                if (parsed && parsed.apiKey && parsed.projectId) {
                    return parsed;
                }
            }
        } catch (e) {
            console.warn("Không thể đọc cấu hình Firebase từ localStorage:", e);
        }
        return DEFAULT_FIREBASE_CONFIG;
    }

    function saveFirebaseConfig(configObj) {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(configObj));
            return true;
        } catch (e) {
            console.error("Lỗi khi lưu cấu hình Firebase:", e);
            return false;
        }
    }

    function clearSavedFirebaseConfig() {
        localStorage.removeItem(STORAGE_KEY);
    }

    function isFirebaseConfigured() {
        const cfg = getFirebaseConfig();
        return !!(cfg && cfg.apiKey && cfg.projectId && cfg.apiKey.trim() !== "" && cfg.projectId.trim() !== "");
    }

    function isAdminEmail(email) {
        if (!email) return false;
        const normalized = email.trim().toLowerCase();
        return DEFAULT_ADMIN_EMAILS.some(adminEmail => adminEmail.toLowerCase() === normalized);
    }

    window.FirebaseConfig = {
        DEFAULT_ADMIN_EMAILS,
        DEFAULT_FIREBASE_CONFIG,
        getFirebaseConfig,
        saveFirebaseConfig,
        clearSavedFirebaseConfig,
        isFirebaseConfigured,
        isAdminEmail
    };
})(window);
