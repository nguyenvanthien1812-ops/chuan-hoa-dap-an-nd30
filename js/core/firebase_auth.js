/**
 * firebase_auth.js - Authentication & License Gatekeeper for DocNormalizer Pro
 */

(function(window) {
    let auth = null;
    let db = null;
    let currentUser = null;
    let currentUserProfile = null;
    let userDocUnsubscribe = null;

    // Callbacks to notify app state changes
    const listeners = [];

    function onUserStateChanged(callback) {
        if (typeof callback === 'function') {
            listeners.push(callback);
        }
    }

    function notifyListeners() {
        const state = getAuthState();
        listeners.forEach(cb => {
            try { cb(state); } catch (e) { console.error("Error in auth listener:", e); }
        });
    }

    function initFirebase() {
        if (!window.firebase) {
            console.warn("Firebase SDK chưa được tải vào trang.");
            return false;
        }

        const config = window.FirebaseConfig.getFirebaseConfig();
        if (!config || !config.apiKey || !config.projectId) {
            console.warn("Chưa thiết lập cấu hình Firebase (apiKey / projectId).");
            return false;
        }

        try {
            if (!firebase.apps.length) {
                firebase.initializeApp(config);
            }
            auth = firebase.auth();
            db = firebase.firestore();

            // Listen to Auth State
            auth.onAuthStateChanged(async (user) => {
                currentUser = user;
                if (user) {
                    await handleUserSignIn(user);
                } else {
                    handleUserSignOut();
                }
            });

            return true;
        } catch (e) {
            console.error("Lỗi khởi tạo Firebase:", e);
            return false;
        }
    }

    async function handleUserSignIn(user) {
        if (!db) return;

        const userRef = db.collection("users").doc(user.uid);
        const isAdmin = window.FirebaseConfig.isAdminEmail(user.email);

        try {
            const docSnap = await userRef.get();
            const now = new Date().toISOString();

            if (!docSnap.exists) {
                // First-time registration
                const newProfile = {
                    uid: user.uid,
                    email: user.email || "",
                    displayName: user.displayName || "Người dùng",
                    photoURL: user.photoURL || "",
                    role: isAdmin ? "admin" : "user",
                    status: isAdmin ? "active" : "pending",
                    plan: isAdmin ? "lifetime" : "none",
                    createdAt: now,
                    lastLoginAt: now,
                    approvedAt: isAdmin ? now : null,
                    expireAt: null
                };
                await userRef.set(newProfile);
                currentUserProfile = newProfile;
            } else {
                // Existing user
                const existing = docSnap.data();
                const updates = {
                    lastLoginAt: now,
                    displayName: user.displayName || existing.displayName || "",
                    photoURL: user.photoURL || existing.photoURL || ""
                };

                // Auto elevate to admin & lifetime if in Admin list
                if (isAdmin && (existing.role !== "admin" || existing.status !== "active" || existing.plan !== "lifetime")) {
                    updates.role = "admin";
                    updates.status = "active";
                    updates.plan = "lifetime";
                    updates.approvedAt = updates.approvedAt || now;
                    updates.expireAt = null;
                }

                await userRef.update(updates);
                currentUserProfile = { ...existing, ...updates };
            }

            // Unsubscribe previous real-time listener if exists
            if (userDocUnsubscribe) {
                userDocUnsubscribe();
            }

            // Listen to real-time changes on user's doc (e.g. when Admin approves)
            userDocUnsubscribe = userRef.onSnapshot((doc) => {
                if (doc.exists) {
                    currentUserProfile = doc.data();
                    renderHeaderAuthUI();
                    renderLicenseGateBanner();
                    notifyListeners();
                }
            });

        } catch (e) {
            console.error("Lỗi đồng bộ thông tin User lên Firestore:", e);
            // Fallback profile if firestore read/write fails
            currentUserProfile = {
                uid: user.uid,
                email: user.email,
                displayName: user.displayName,
                photoURL: user.photoURL,
                role: isAdmin ? "admin" : "user",
                status: isAdmin ? "active" : "pending",
                plan: isAdmin ? "lifetime" : "none"
            };
        }

        renderHeaderAuthUI();
        renderLicenseGateBanner();
        notifyListeners();
    }

    function handleUserSignOut() {
        if (userDocUnsubscribe) {
            userDocUnsubscribe();
            userDocUnsubscribe = null;
        }
        currentUser = null;
        currentUserProfile = null;
        renderHeaderAuthUI();
        renderLicenseGateBanner();
        notifyListeners();
    }

    async function loginWithGoogle() {
        if (!auth) {
            const ok = initFirebase();
            if (!ok) {
                alert("Lỗi kết nối Firebase Authentication. Vui lòng kiểm tra lại kết nối mạng hoặc cấu hình Firebase!");
                return;
            }
        }

        const provider = new firebase.auth.GoogleAuthProvider();
        provider.setCustomParameters({ prompt: 'select_account' });

        try {
            await auth.signInWithPopup(provider);
        } catch (error) {
            console.warn("Popup sign-in gặp lỗi hoặc bị chặn, chuyển sang redirect:", error);
            if (error.code === 'auth/popup-blocked' || error.code === 'auth/popup-closed-by-user') {
                try {
                    await auth.signInWithRedirect(provider);
                } catch (redErr) {
                    console.error("Lỗi đăng nhập Google Redirect:", redErr);
                    alert("Không thể đăng nhập Google: " + redErr.message);
                }
            } else {
                alert("Đăng nhập thất bại: " + error.message);
            }
        }
    }

    async function logout() {
        if (!auth) return;
        try {
            await auth.signOut();
            if (window.showToast) {
                window.showToast("Đã đăng xuất tài khoản!", "info");
            }
        } catch (e) {
            console.error("Lỗi đăng xuất:", e);
        }
    }

    function getAuthState() {
        const check = checkLicenseAccess();
        return {
            isConfigured: window.FirebaseConfig.isFirebaseConfigured(),
            isLoggedIn: !!currentUser,
            user: currentUser,
            profile: currentUserProfile,
            isAdmin: currentUserProfile && currentUserProfile.role === 'admin',
            license: check
        };
    }

    function checkLicenseAccess() {
        // If Firebase is not yet configured, allow offline / dev mode with notice
        if (!window.FirebaseConfig.isFirebaseConfigured()) {
            return {
                allowed: true,
                status: 'unconfigured',
                plan: 'offline',
                message: 'Chế độ ngoại tuyến (Chưa kích hoạt Firebase Auth).'
            };
        }

        if (!currentUser) {
            return {
                allowed: false,
                reason: 'not_logged_in',
                status: 'anonymous',
                message: 'Vui lòng đăng nhập bằng tài khoản Google để sử dụng tính năng này.'
            };
        }

        if (!currentUserProfile) {
            return {
                allowed: false,
                reason: 'loading',
                status: 'loading',
                message: 'Đang tải thông tin bản quyền...'
            };
        }

        const { status, plan, expireAt, role } = currentUserProfile;

        // Admin always has full access
        if (role === 'admin') {
            return {
                allowed: true,
                status: 'active',
                plan: 'lifetime',
                isAdmin: true,
                message: 'Quản trị viên (Bản quyền Vĩnh Viễn)'
            };
        }

        if (status === 'blocked') {
            return {
                allowed: false,
                reason: 'blocked',
                status: 'blocked',
                message: 'Tài khoản của bạn đã bị khóa quyền truy cập. Vui lòng liên hệ Admin để được hỗ trợ.'
            };
        }

        if (status === 'pending') {
            return {
                allowed: false,
                reason: 'pending',
                status: 'pending',
                message: 'Tài khoản của bạn đang chờ Quản trị viên (Admin) phê duyệt bản quyền. Vui lòng liên hệ Admin để kích hoạt.'
            };
        }

        if (status === 'active') {
            if (plan === 'lifetime') {
                return {
                    allowed: true,
                    status: 'active',
                    plan: 'lifetime',
                    message: 'Bản quyền Vĩnh Viễn'
                };
            }

            if (plan === '1_year' || expireAt) {
                if (!expireAt) {
                    return {
                        allowed: true,
                        status: 'active',
                        plan: '1_year',
                        message: 'Bản quyền 1 Năm'
                    };
                }

                const expireTime = new Date(expireAt).getTime();
                const nowTime = new Date().getTime();

                if (nowTime > expireTime) {
                    const expireDateStr = new Date(expireAt).toLocaleDateString('vi-VN');
                    return {
                        allowed: false,
                        reason: 'expired',
                        status: 'expired',
                        expireAt,
                        message: `Bản quyền 1 Năm của bạn đã hết hạn vào ngày ${expireDateStr}. Vui lòng liên hệ Admin để gia hạn.`
                    };
                }

                const daysLeft = Math.ceil((expireTime - nowTime) / (1000 * 60 * 60 * 24));
                const expireDateStr = new Date(expireAt).toLocaleDateString('vi-VN');
                return {
                    allowed: true,
                    status: 'active',
                    plan: '1_year',
                    expireAt,
                    daysLeft,
                    message: `Hạn dùng đến ngày ${expireDateStr} (Còn ${daysLeft} ngày)`
                };
            }
        }

        return {
            allowed: false,
            reason: 'unknown',
            status: status || 'pending',
            message: 'Tài khoản chưa có giấy phép hợp lệ.'
        };
    }

    function renderHeaderAuthUI() {
        const container = document.getElementById("headerAuthContainer");
        if (!container) return;

        const isCfg = window.FirebaseConfig.isFirebaseConfigured();

        if (!currentUser) {
            container.innerHTML = `
                <button id="btnLoginGoogle" class="btn-google-login" title="Đăng nhập tài khoản Google">
                    <i class="fa-brands fa-google"></i>
                    <span>Đăng nhập</span>
                </button>
            `;

            const btnLogin = document.getElementById("btnLoginGoogle");
            if (btnLogin) btnLogin.onclick = loginWithGoogle;
            return;
        }

        // Logged in
        const profile = currentUserProfile || {};
        const displayName = profile.displayName || currentUser.displayName || "Người dùng";
        const photoURL = profile.photoURL || currentUser.photoURL || "https://www.gravatar.com/avatar/?d=mp";
        const isAdmin = profile.role === 'admin';
        const license = checkLicenseAccess();

        let badgeHtml = '';
        if (isAdmin || license.plan === 'lifetime') {
            badgeHtml = `<span class="license-badge lifetime"><i class="fa-solid fa-crown"></i> Vĩnh viễn</span>`;
        } else if (license.status === 'active' && license.plan === '1_year') {
            badgeHtml = `<span class="license-badge yearly" title="${license.message}"><i class="fa-solid fa-calendar-check"></i> ${license.daysLeft ? license.daysLeft + ' ngày' : '1 Năm'}</span>`;
        } else if (license.status === 'pending') {
            badgeHtml = `<span class="license-badge pending"><i class="fa-solid fa-clock"></i> Chờ duyệt</span>`;
        } else {
            badgeHtml = `<span class="license-badge expired"><i class="fa-solid fa-lock"></i> Hết hạn</span>`;
        }

        container.innerHTML = `
            <div class="user-profile-pill">
                <img src="${photoURL}" alt="${displayName}" class="user-avatar" referrerpolicy="no-referrer">
                <div class="user-info-text">
                    <span class="user-display-name">${displayName}</span>
                    ${badgeHtml}
                </div>
            </div>

            ${isAdmin ? `
                <button id="btnOpenAdminPanel" class="btn-admin-panel" title="Mở bảng điều khiển Quản trị viên">
                    <i class="fa-solid fa-user-shield"></i>
                    <span>Quản trị</span>
                </button>
            ` : ''}

            <button id="btnLogoutUser" class="btn-logout" title="Đăng xuất">
                <i class="fa-solid fa-right-from-bracket"></i>
            </button>
        `;

        const btnAdmin = document.getElementById("btnOpenAdminPanel");
        if (btnAdmin && window.AdminPanel) {
            btnAdmin.onclick = () => window.AdminPanel.openDashboard();
        }

        const btnLogout = document.getElementById("btnLogoutUser");
        if (btnLogout) {
            btnLogout.onclick = logout;
        }
    }

    function renderLicenseGateBanner() {
        const bannerContainer = document.getElementById("licenseGateContainer");
        if (!bannerContainer) return;

        const license = checkLicenseAccess();

        if (license.allowed) {
            bannerContainer.innerHTML = "";
            bannerContainer.style.display = "none";
            return;
        }

        // Not allowed -> Render warning banner
        bannerContainer.style.display = "block";
        const isBlocked = license.status === 'blocked' || license.status === 'expired';

        bannerContainer.innerHTML = `
            <div class="license-gate-banner ${isBlocked ? 'blocked' : ''}">
                <div class="license-gate-content">
                    <div class="license-gate-icon">
                        <i class="fa-solid ${isBlocked ? 'fa-triangle-exclamation' : 'fa-hourglass-half'}"></i>
                    </div>
                    <div>
                        <div class="license-gate-title">
                            ${license.status === 'pending' ? 'Tài khoản của bạn đang chờ phê duyệt bản quyền' : 
                              license.status === 'expired' ? 'Bản quyền sử dụng của bạn đã hết hạn' : 
                              license.status === 'anonymous' ? 'Vui lòng đăng nhập để sử dụng DocNormalizer Pro' : 
                              'Tài khoản bị tạm khóa quyền sử dụng'}
                        </div>
                        <div class="license-gate-desc">
                            ${license.message} Liên hệ Quản trị viên (Admin) qua Zalo/SĐT để được cấp quyền dùng 1 Năm hoặc Vĩnh Viễn.
                        </div>
                    </div>
                </div>
                <div class="license-gate-actions">
                    ${license.status === 'anonymous' ? `
                        <button class="btn btn-primary" onclick="window.FirebaseAuth.loginWithGoogle()">
                            <i class="fa-brands fa-google"></i> Đăng nhập ngay
                        </button>
                    ` : `
                        <a href="https://zalo.me/0393732038" target="_blank" class="btn-contact-admin">
                            <i class="fa-solid fa-comment-dots"></i> Liên hệ Admin kích hoạt
                        </a>
                    `}
                </div>
            </div>
        `;
    }

    // Auto init on DOMContentLoaded
    document.addEventListener("DOMContentLoaded", () => {
        initFirebase();
        renderHeaderAuthUI();
        renderLicenseGateBanner();
    });

    window.FirebaseAuth = {
        initFirebase,
        loginWithGoogle,
        logout,
        getAuthState,
        checkLicenseAccess,
        onUserStateChanged,
        renderHeaderAuthUI,
        renderLicenseGateBanner,
        getDb: () => db
    };
})(window);
