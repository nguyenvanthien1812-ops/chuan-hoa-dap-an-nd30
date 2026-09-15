/**
 * admin_panel.js - Bảng Quản trị viên & Quản lý Phê duyệt Bản quyền cho DocNormalizer Pro
 */

(function(window) {
    let usersList = [];
    let unsubscribeUsers = null;
    let currentFilter = 'all';
    let searchQuery = '';

    function openDashboard() {
        const modal = document.getElementById("adminDashboardModal");
        if (!modal) return;

        const authState = window.FirebaseAuth ? window.FirebaseAuth.getAuthState() : null;
        if (!authState || !authState.isAdmin) {
            alert("Bạn không có quyền Quản trị viên (Admin).");
            return;
        }

        modal.classList.add("show");
        subscribeUsersList();
    }

    function closeDashboard() {
        const modal = document.getElementById("adminDashboardModal");
        if (modal) {
            modal.classList.remove("show");
        }
        if (unsubscribeUsers) {
            unsubscribeUsers();
            unsubscribeUsers = null;
        }
    }

    function subscribeUsersList() {
        const db = window.FirebaseAuth ? window.FirebaseAuth.getDb() : null;
        if (!db) {
            console.error("Chưa kết nối được Firestore DB.");
            return;
        }

        if (unsubscribeUsers) {
            unsubscribeUsers();
        }

        const tableBody = document.getElementById("adminUsersTableBody");
        if (tableBody) {
            tableBody.innerHTML = `<tr><td colspan="6" style="text-align: center; padding: 24px;"><i class="fa-solid fa-spinner fa-spin"></i> Đang tải danh sách người dùng...</td></tr>`;
        }

        try {
            unsubscribeUsers = db.collection("users")
                .onSnapshot((snapshot) => {
                    usersList = [];
                    snapshot.forEach(doc => {
                        const data = doc.data();
                        if (data) {
                            usersList.push({ ...data, uid: data.uid || doc.id });
                        }
                    });
                    // Sắp xếp theo ngày tạo mới nhất (client-side)
                    usersList.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
                    renderUserStats();
                    renderUsersTable();
                }, (error) => {
                    console.error("Lỗi khi tải danh sách users từ Firestore:", error);
                    let helpMsg = '';
                    if (error.code === 'permission-denied' || (error.message && error.message.includes('permission'))) {
                        helpMsg = `
                            <div style="margin-top: 10px; font-size: 13px; line-height: 1.5; color: var(--text-secondary);">
                                <strong>👉 Nguyên nhân:</strong> Firestore Security Rules trên Firebase Console đang chặn quyền đọc dữ liệu.<br>
                                <strong>👉 Cách khắc phục:</strong> Vào <strong>Firebase Console &rarr; Firestore Database &rarr; Tab Rules</strong>, dán:
                                <pre style="background: rgba(0,0,0,0.3); padding: 8px; border-radius: 6px; margin: 8px 0; font-family: monospace; color: #10b981; text-align: left;">rules_version = '2';\nservice cloud.firestore {\n  match /databases/{database}/documents {\n    match /{document=**} {\n      allow read, write: if request.auth != null;\n    }\n  }\n}</pre>
                                rồi bấm nút <strong>Publish</strong>.
                            </div>
                        `;
                    }
                    if (tableBody) {
                        tableBody.innerHTML = `
                            <tr>
                                <td colspan="6" style="text-align: center; color: var(--accent-rose); padding: 24px;">
                                    <div><i class="fa-solid fa-triangle-exclamation"></i> <strong>Lỗi tải dữ liệu:</strong> ${error.message}</div>
                                    ${helpMsg}
                                </td>
                            </tr>
                        `;
                    }
                });
        } catch (e) {
            console.error("Lỗi truy vấn collection users:", e);
        }
    }

    function renderUserStats() {
        let total = usersList.length;
        let pending = 0;
        let yearly = 0;
        let lifetime = 0;

        usersList.forEach(u => {
            if (u.role === 'admin' || u.plan === 'lifetime') {
                lifetime++;
            } else if (u.status === 'pending') {
                pending++;
            } else if (u.status === 'active' && u.plan === '1_year') {
                yearly++;
            }
        });

        const elTotal = document.getElementById("statTotalUsers");
        const elPending = document.getElementById("statPendingUsers");
        const elYearly = document.getElementById("statYearlyUsers");
        const elLifetime = document.getElementById("statLifetimeUsers");

        if (elTotal) elTotal.innerText = total;
        if (elPending) elPending.innerText = pending;
        if (elYearly) elYearly.innerText = yearly;
        if (elLifetime) elLifetime.innerText = lifetime;
    }

    function filterUsers() {
        return usersList.filter(u => {
            // Check status / plan filter
            if (currentFilter === 'pending' && u.status !== 'pending') return false;
            if (currentFilter === '1_year' && (u.status !== 'active' || u.plan !== '1_year')) return false;
            if (currentFilter === 'lifetime' && u.plan !== 'lifetime' && u.role !== 'admin') return false;
            if (currentFilter === 'blocked' && u.status !== 'blocked') return false;

            // Check search text (email or displayName)
            if (searchQuery.trim() !== '') {
                const q = searchQuery.toLowerCase();
                const name = (u.displayName || '').toLowerCase();
                const email = (u.email || '').toLowerCase();
                if (!name.includes(q) && !email.includes(q)) return false;
            }

            return true;
        });
    }

    function renderUsersTable() {
        const tableBody = document.getElementById("adminUsersTableBody");
        if (!tableBody) return;

        const filtered = filterUsers();

        if (filtered.length === 0) {
            tableBody.innerHTML = `
                <tr>
                    <td colspan="6" style="text-align: center; padding: 24px; color: var(--text-muted);">
                        Không tìm thấy người dùng nào phù hợp.
                    </td>
                </tr>
            `;
            return;
        }

        const now = new Date().getTime();

        tableBody.innerHTML = filtered.map(user => {
            const isSuperAdmin = window.FirebaseConfig.isAdminEmail(user.email);
            const photoURL = user.photoURL || "https://www.gravatar.com/avatar/?d=mp";
            const displayName = escapeHtml(user.displayName || "Không tên");
            const email = escapeHtml(user.email || "Chưa có email");
            
            // Format dates
            const createdDate = user.createdAt ? new Date(user.createdAt).toLocaleDateString('vi-VN') : '—';
            let expireDateStr = '—';
            let isExpired = false;

            if (user.role === 'admin' || user.plan === 'lifetime') {
                expireDateStr = '<span class="text-success font-weight-bold"><i class="fa-solid fa-infinity"></i> Vô thời hạn</span>';
            } else if (user.expireAt) {
                const expTime = new Date(user.expireAt).getTime();
                isExpired = now > expTime;
                const formatted = new Date(user.expireAt).toLocaleDateString('vi-VN');
                expireDateStr = isExpired 
                    ? `<span class="text-danger font-weight-bold">${formatted} (Đã hết hạn)</span>`
                    : `<span class="text-primary">${formatted}</span>`;
            }

            // Status Badge
            let statusBadge = '';
            if (user.role === 'admin') {
                statusBadge = '<span class="status-badge admin"><i class="fa-solid fa-crown"></i> Admin</span>';
            } else if (user.status === 'blocked') {
                statusBadge = '<span class="status-badge blocked"><i class="fa-solid fa-ban"></i> Đã khóa</span>';
            } else if (user.status === 'pending') {
                statusBadge = '<span class="status-badge pending"><i class="fa-solid fa-hourglass-half"></i> Chờ duyệt</span>';
            } else if (isExpired) {
                statusBadge = '<span class="status-badge expired"><i class="fa-solid fa-clock"></i> Hết hạn</span>';
            } else {
                statusBadge = '<span class="status-badge active"><i class="fa-solid fa-circle-check"></i> Đang hoạt động</span>';
            }

            // Plan Badge
            let planBadge = '';
            if (user.role === 'admin' || user.plan === 'lifetime') {
                planBadge = '<span class="license-badge lifetime"><i class="fa-solid fa-gem"></i> Vĩnh viễn</span>';
            } else if (user.plan === '1_year') {
                planBadge = '<span class="license-badge yearly"><i class="fa-solid fa-calendar-check"></i> 1 Năm</span>';
            } else {
                planBadge = '<span class="license-badge pending"><i class="fa-solid fa-minus"></i> Chưa cấp</span>';
            }

            // Actions Buttons
            let actionButtons = '';
            if (isSuperAdmin) {
                actionButtons = `<span style="font-size: 12px; color: var(--text-muted); font-style: italic;"><i class="fa-solid fa-shield-halved"></i> Quản trị viên tối cao</span>`;
            } else {
                actionButtons = `
                    <div style="display: flex; gap: 4px; flex-wrap: wrap;">
                        <button class="btn-action-sm btn-approve-year" onclick="window.AdminPanel.approveOneYear('${user.uid}')" title="Cấp bản quyền 1 Năm (365 ngày)">
                            <i class="fa-solid fa-calendar-plus"></i> 1 Năm
                        </button>
                        <button class="btn-action-sm btn-approve-lifetime" onclick="window.AdminPanel.approveLifetime('${user.uid}')" title="Cấp bản quyền Vĩnh Viễn">
                            <i class="fa-solid fa-crown"></i> Vĩnh viễn
                        </button>
                        ${user.plan === '1_year' ? `
                            <button class="btn-action-sm btn-approve-year" onclick="window.AdminPanel.extendOneYear('${user.uid}', '${user.expireAt || ''}')" title="Gia hạn thêm 1 Năm">
                                <i class="fa-solid fa-arrows-rotate"></i> Gia hạn
                            </button>
                        ` : ''}
                        <button class="btn-action-sm ${user.status === 'blocked' ? 'btn-approve-year' : 'btn-block'}" onclick="window.AdminPanel.toggleBlock('${user.uid}', '${user.status}')" title="${user.status === 'blocked' ? 'Mở khóa tài khoản' : 'Khóa tài khoản này'}">
                            <i class="fa-solid ${user.status === 'blocked' ? 'fa-unlock' : 'fa-user-slash'}"></i>
                        </button>
                        <button class="btn-action-sm btn-block" onclick="window.AdminPanel.deleteUser('${user.uid}', '${user.email}')" title="Xóa tài khoản khỏi danh sách">
                            <i class="fa-solid fa-trash-can"></i>
                        </button>
                    </div>
                `;
            }

            return `
                <tr>
                    <td>
                        <div style="display: flex; align-items: center; gap: 10px;">
                            <img src="${photoURL}" alt="${displayName}" class="user-avatar" referrerpolicy="no-referrer">
                            <div>
                                <strong style="color: var(--text-dark);">${displayName}</strong>
                                <div style="font-size: 11px; color: var(--text-muted); font-family: monospace;">${user.uid.substring(0, 10)}...</div>
                            </div>
                        </div>
                    </td>
                    <td><span style="font-family: monospace; font-size: 13px;">${email}</span></td>
                    <td>${statusBadge}</td>
                    <td>${planBadge}</td>
                    <td>${expireDateStr}</td>
                    <td>${actionButtons}</td>
                </tr>
            `;
        }).join('');
    }

    async function approveOneYear(uid) {
        const db = window.FirebaseAuth ? window.FirebaseAuth.getDb() : null;
        if (!db) return;

        const now = new Date();
        const oneYearLater = new Date(now.getTime() + 365 * 24 * 60 * 60 * 1000);

        try {
            await db.collection("users").doc(uid).update({
                status: 'active',
                plan: '1_year',
                approvedAt: now.toISOString(),
                expireAt: oneYearLater.toISOString()
            });
            if (window.showToast) {
                window.showToast("Đã duyệt bản quyền 1 Năm thành công!", "success");
            }
        } catch (e) {
            console.error("Lỗi duyệt 1 năm:", e);
            alert("Lỗi khi duyệt bản quyền 1 năm: " + e.message);
        }
    }

    async function approveLifetime(uid) {
        const db = window.FirebaseAuth ? window.FirebaseAuth.getDb() : null;
        if (!db) return;

        const now = new Date();

        try {
            await db.collection("users").doc(uid).update({
                status: 'active',
                plan: 'lifetime',
                approvedAt: now.toISOString(),
                expireAt: null
            });
            if (window.showToast) {
                window.showToast("Đã duyệt bản quyền Vĩnh Viễn thành công!", "success");
            }
        } catch (e) {
            console.error("Lỗi duyệt vĩnh viễn:", e);
            alert("Lỗi khi duyệt bản quyền vĩnh viễn: " + e.message);
        }
    }

    async function extendOneYear(uid, currentExpireAt) {
        const db = window.FirebaseAuth ? window.FirebaseAuth.getDb() : null;
        if (!db) return;

        let baseDate = new Date();
        if (currentExpireAt) {
            const expDate = new Date(currentExpireAt);
            if (expDate.getTime() > baseDate.getTime()) {
                baseDate = expDate;
            }
        }

        const newExpire = new Date(baseDate.getTime() + 365 * 24 * 60 * 60 * 1000);

        try {
            await db.collection("users").doc(uid).update({
                status: 'active',
                plan: '1_year',
                expireAt: newExpire.toISOString()
            });
            if (window.showToast) {
                window.showToast(`Đã gia hạn thêm 1 Năm (Đến ${newExpire.toLocaleDateString('vi-VN')})!`, "success");
            }
        } catch (e) {
            console.error("Lỗi gia hạn:", e);
            alert("Lỗi khi gia hạn: " + e.message);
        }
    }

    async function toggleBlock(uid, currentStatus) {
        const db = window.FirebaseAuth ? window.FirebaseAuth.getDb() : null;
        if (!db) return;

        const newStatus = currentStatus === 'blocked' ? 'active' : 'blocked';
        const confirmMsg = newStatus === 'blocked' 
            ? "Bạn có chắc chắn muốn KHÓA tài khoản này không?" 
            : "Bạn có muốn MỞ KHÓA tài khoản này không?";

        if (!confirm(confirmMsg)) return;

        try {
            await db.collection("users").doc(uid).update({
                status: newStatus
            });
            if (window.showToast) {
                window.showToast(newStatus === 'blocked' ? "Đã khóa tài khoản!" : "Đã mở khóa tài khoản!", "info");
            }
        } catch (e) {
            console.error("Lỗi đổi trạng thái block:", e);
            alert("Lỗi: " + e.message);
        }
    }

    async function deleteUser(uid, email) {
        const db = window.FirebaseAuth ? window.FirebaseAuth.getDb() : null;
        if (!db) return;

        if (!confirm(`Bạn có chắc muốn xóa tài khoản [${email}] khỏi hệ thống?`)) return;

        try {
            await db.collection("users").doc(uid).delete();
            if (window.showToast) {
                window.showToast("Đã xóa tài khoản khỏi danh sách!", "info");
            }
        } catch (e) {
            console.error("Lỗi xóa user:", e);
            alert("Lỗi khi xóa người dùng: " + e.message);
        }
    }

    // ================== FIREBASE CONFIG MODAL ==================

    function showConfigModal(noticeText) {
        const modal = document.getElementById("firebaseConfigModal");
        if (!modal) return;

        const cfg = window.FirebaseConfig.getFirebaseConfig();
        const elApiKey = document.getElementById("fbCfgApiKey");
        const elAuthDomain = document.getElementById("fbCfgAuthDomain");
        const elProjectId = document.getElementById("fbCfgProjectId");
        const elStorageBucket = document.getElementById("fbCfgStorageBucket");
        const elMessagingSenderId = document.getElementById("fbCfgMessagingSenderId");
        const elAppId = document.getElementById("fbCfgAppId");

        if (elApiKey) elApiKey.value = cfg.apiKey || "";
        if (elAuthDomain) elAuthDomain.value = cfg.authDomain || "";
        if (elProjectId) elProjectId.value = cfg.projectId || "";
        if (elStorageBucket) elStorageBucket.value = cfg.storageBucket || "";
        if (elMessagingSenderId) elMessagingSenderId.value = cfg.messagingSenderId || "";
        if (elAppId) elAppId.value = cfg.appId || "";

        const noticeEl = document.getElementById("fbCfgNotice");
        if (noticeEl) {
            if (noticeText) {
                noticeEl.innerText = noticeText;
                noticeEl.style.display = "block";
            } else {
                noticeEl.style.display = "none";
            }
        }

        modal.classList.add("show");
    }

    function closeConfigModal() {
        const modal = document.getElementById("firebaseConfigModal");
        if (modal) {
            modal.classList.remove("show");
        }
    }

    function saveConfigFromForm() {
        const apiKey = (document.getElementById("fbCfgApiKey")?.value || "").trim();
        const authDomain = (document.getElementById("fbCfgAuthDomain")?.value || "").trim();
        const projectId = (document.getElementById("fbCfgProjectId")?.value || "").trim();
        const storageBucket = (document.getElementById("fbCfgStorageBucket")?.value || "").trim();
        const messagingSenderId = (document.getElementById("fbCfgMessagingSenderId")?.value || "").trim();
        const appId = (document.getElementById("fbCfgAppId")?.value || "").trim();

        if (!apiKey || !projectId) {
            alert("Vui lòng nhập ít nhất apiKey và projectId!");
            return;
        }

        const newConfig = {
            apiKey,
            authDomain: authDomain || `${projectId}.firebaseapp.com`,
            projectId,
            storageBucket: storageBucket || `${projectId}.firebasestorage.app`,
            messagingSenderId,
            appId
        };

        const success = window.FirebaseConfig.saveFirebaseConfig(newConfig);
        if (success) {
            alert("Đã lưu cấu hình Firebase thành công! Trang web sẽ tự tải lại để kích hoạt.");
            closeConfigModal();
            window.location.reload();
        } else {
            alert("Lỗi khi lưu cấu hình vào bộ nhớ trình duyệt.");
        }
    }

    function clearConfig() {
        if (confirm("Bạn có chắc muốn xóa cấu hình Firebase đã lưu trên trình duyệt này?")) {
            window.FirebaseConfig.clearSavedFirebaseConfig();
            alert("Đã xóa cấu hình Firebase. Trang sẽ tải lại.");
            window.location.reload();
        }
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    // Set up event listeners for filters & search
    document.addEventListener("DOMContentLoaded", () => {
        const searchInput = document.getElementById("adminUserSearch");
        if (searchInput) {
            searchInput.addEventListener("input", (e) => {
                searchQuery = e.target.value;
                renderUsersTable();
            });
        }

        const filterBtns = document.querySelectorAll(".admin-filter-btn");
        filterBtns.forEach(btn => {
            btn.addEventListener("click", () => {
                filterBtns.forEach(b => b.classList.remove("active"));
                btn.classList.add("active");
                currentFilter = btn.getAttribute("data-filter") || 'all';
                renderUsersTable();
            });
        });
    });

    window.AdminPanel = {
        openDashboard,
        closeDashboard,
        approveOneYear,
        approveLifetime,
        extendOneYear,
        toggleBlock,
        deleteUser,
        showConfigModal,
        closeConfigModal,
        saveConfigFromForm,
        clearConfig
    };

})(window);
