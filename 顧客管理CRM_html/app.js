// ===== CRM Application - Main JavaScript =====

// ===== Data Management =====
class CRMApp {
    constructor() {
        this.customers = [];
        this.history = [];
        this.currentCustomerId = null;
        this.init();
    }

    init() {
        this.loadData();
        this.bindEvents();
        this.renderDashboard();
        this.renderCustomersTable();
        this.renderHistoryList();
        this.updateStats();
    }

    // ===== LocalStorage =====
    loadData() {
        const savedCustomers = localStorage.getItem('crm_customers');
        const savedHistory = localStorage.getItem('crm_history');
        
        if (savedCustomers) {
            this.customers = JSON.parse(savedCustomers);
        }
        if (savedHistory) {
            this.history = JSON.parse(savedHistory);
        }
    }

    saveData() {
        localStorage.setItem('crm_customers', JSON.stringify(this.customers));
        localStorage.setItem('crm_history', JSON.stringify(this.history));
    }

    // ===== Event Bindings =====
    bindEvents() {
        // Navigation
        document.querySelectorAll('.nav-item').forEach(item => {
            item.addEventListener('click', (e) => {
                e.preventDefault();
                const view = item.dataset.view;
                this.switchView(view);
            });
        });

        // Search
        document.getElementById('globalSearch').addEventListener('input', (e) => {
            this.handleSearch(e.target.value);
        });

        // Add customer buttons
        document.getElementById('addCustomerBtn').addEventListener('click', () => {
            this.openCustomerModal();
        });
        document.getElementById('addFirstCustomer')?.addEventListener('click', () => {
            this.openCustomerModal();
        });

        // Customer form
        document.getElementById('customerForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveCustomer();
        });

        // Customer modal close
        document.getElementById('closeCustomerModal').addEventListener('click', () => {
            this.closeCustomerModal();
        });
        document.getElementById('cancelCustomer').addEventListener('click', () => {
            this.closeCustomerModal();
        });
        document.querySelector('#customerModal .modal-overlay').addEventListener('click', () => {
            this.closeCustomerModal();
        });

        // Detail modal
        document.getElementById('closeDetailModal').addEventListener('click', () => {
            this.closeDetailModal();
        });
        document.querySelector('#customerDetailModal .modal-overlay').addEventListener('click', () => {
            this.closeDetailModal();
        });

        // Detail modal tabs
        document.querySelectorAll('.detail-tabs .tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this.switchDetailTab(btn.dataset.tab);
            });
        });

        // Edit/Delete customer
        document.getElementById('editCustomerBtn').addEventListener('click', () => {
            this.editCurrentCustomer();
        });
        document.getElementById('deleteCustomerBtn').addEventListener('click', () => {
            this.confirmDeleteCustomer();
        });

        // Add history
        document.getElementById('addHistoryBtn').addEventListener('click', () => {
            this.openHistoryModal();
        });

        // History form
        document.getElementById('historyForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.saveHistory();
        });

        // History modal close
        document.getElementById('closeHistoryModal').addEventListener('click', () => {
            this.closeHistoryModal();
        });
        document.getElementById('cancelHistory').addEventListener('click', () => {
            this.closeHistoryModal();
        });
        document.querySelector('#historyModal .modal-overlay').addEventListener('click', () => {
            this.closeHistoryModal();
        });

        // Confirm modal
        document.getElementById('confirmCancel').addEventListener('click', () => {
            this.closeConfirmModal();
        });
        document.querySelector('#confirmModal .modal-overlay').addEventListener('click', () => {
            this.closeConfirmModal();
        });

        // Filters
        document.getElementById('statusFilter').addEventListener('change', () => {
            this.renderCustomersTable();
        });
        document.getElementById('rankFilter').addEventListener('change', () => {
            this.renderCustomersTable();
        });
        document.getElementById('sortFilter').addEventListener('change', () => {
            this.renderCustomersTable();
        });
        document.getElementById('historyTypeFilter').addEventListener('change', () => {
            this.renderHistoryList();
        });
        document.getElementById('historyPeriodFilter').addEventListener('change', () => {
            this.renderHistoryList();
        });
    }

    // ===== Navigation =====
    switchView(viewName) {
        document.querySelectorAll('.nav-item').forEach(item => {
            item.classList.toggle('active', item.dataset.view === viewName);
        });
        document.querySelectorAll('.view').forEach(view => {
            view.classList.toggle('active', view.id === viewName + 'View');
        });
    }

    // ===== Search =====
    handleSearch(query) {
        const currentView = document.querySelector('.view.active').id;
        
        if (currentView === 'customersView') {
            this.renderCustomersTable(query);
        } else if (currentView === 'historyView') {
            this.renderHistoryList(query);
        }
    }

    // ===== Dashboard =====
    updateStats() {
        document.getElementById('totalCustomers').textContent = this.customers.length;
        document.getElementById('activeCustomers').textContent = 
            this.customers.filter(c => c.status === 'active').length;
        document.getElementById('totalHistory').textContent = this.history.length;
        
        // Monthly history
        const now = new Date();
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        document.getElementById('monthlyHistory').textContent = 
            this.history.filter(h => new Date(h.date) >= monthStart).length;
        
        // Today's count
        const today = new Date().toDateString();
        document.getElementById('todayCount').textContent = 
            this.history.filter(h => new Date(h.date).toDateString() === today).length;
    }

    renderDashboard() {
        this.renderRecentCustomers();
        this.renderRecentHistory();
    }

    renderRecentCustomers() {
        const container = document.getElementById('recentCustomers');
        const recent = [...this.customers]
            .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
            .slice(0, 5);
        
        if (recent.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <p>顧客データがありません</p>
                </div>
            `;
            return;
        }

        container.innerHTML = recent.map(customer => `
            <div class="recent-item" data-id="${customer.id}">
                <div class="recent-avatar">${this.getInitials(customer.name)}</div>
                <div class="recent-info">
                    <div class="recent-name">${this.escapeHtml(customer.name)}</div>
                    <div class="recent-detail">${this.escapeHtml(customer.companyName || '個人')}</div>
                </div>
                <div class="recent-date">${this.formatDate(customer.createdAt)}</div>
            </div>
        `).join('');

        container.querySelectorAll('.recent-item').forEach(item => {
            item.addEventListener('click', () => {
                this.openCustomerDetail(item.dataset.id);
            });
        });
    }

    renderRecentHistory() {
        const container = document.getElementById('recentHistory');
        const recent = [...this.history]
            .sort((a, b) => new Date(b.date) - new Date(a.date))
            .slice(0, 5);
        
        if (recent.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <p>対応履歴がありません</p>
                </div>
            `;
            return;
        }

        container.innerHTML = recent.map(h => {
            const customer = this.customers.find(c => c.id === h.customerId);
            return `
                <div class="recent-item" data-id="${h.customerId}">
                    <div class="recent-avatar">${customer ? this.getInitials(customer.name) : '?'}</div>
                    <div class="recent-info">
                        <div class="recent-name">${this.escapeHtml(h.subject)}</div>
                        <div class="recent-detail">${customer ? this.escapeHtml(customer.name) : '削除済み'}</div>
                    </div>
                    <div class="recent-date">${this.formatDateTime(h.date)}</div>
                </div>
            `;
        }).join('');

        container.querySelectorAll('.recent-item').forEach(item => {
            item.addEventListener('click', () => {
                if (item.dataset.id) {
                    this.openCustomerDetail(item.dataset.id);
                }
            });
        });
    }

    // ===== Customers Table =====
    renderCustomersTable(searchQuery = '') {
        const tbody = document.getElementById('customersTableBody');
        const emptyState = document.getElementById('customersEmptyState');
        const table = document.querySelector('.customers-table');
        
        let filtered = [...this.customers];
        
        // Search filter
        if (searchQuery) {
            const query = searchQuery.toLowerCase();
            filtered = filtered.filter(c => 
                c.name.toLowerCase().includes(query) ||
                (c.companyName && c.companyName.toLowerCase().includes(query)) ||
                (c.phone && c.phone.includes(query)) ||
                (c.email && c.email.toLowerCase().includes(query))
            );
        }
        
        // Status filter
        const statusFilter = document.getElementById('statusFilter').value;
        if (statusFilter) {
            filtered = filtered.filter(c => c.status === statusFilter);
        }
        
        // Rank filter
        const rankFilter = document.getElementById('rankFilter').value;
        if (rankFilter) {
            filtered = filtered.filter(c => c.rank === rankFilter);
        }
        
        // Sort
        const sortFilter = document.getElementById('sortFilter').value;
        switch (sortFilter) {
            case 'newest':
                filtered.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
                break;
            case 'oldest':
                filtered.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
                break;
            case 'name':
                filtered.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
                break;
            case 'company':
                filtered.sort((a, b) => (a.companyName || '').localeCompare(b.companyName || '', 'ja'));
                break;
        }
        
        if (filtered.length === 0) {
            table.style.display = 'none';
            emptyState.style.display = 'flex';
            return;
        }
        
        table.style.display = 'table';
        emptyState.style.display = 'none';
        
        tbody.innerHTML = filtered.map(customer => `
            <tr data-id="${customer.id}">
                <td>
                    <div class="customer-name-cell">
                        <div class="customer-avatar">${this.getInitials(customer.name)}</div>
                        <div>
                            <div style="font-weight: 500;">${this.escapeHtml(customer.name)}</div>
                            ${customer.nameKana ? `<div style="font-size: 0.8rem; color: var(--text-muted);">${this.escapeHtml(customer.nameKana)}</div>` : ''}
                        </div>
                    </div>
                </td>
                <td>${this.escapeHtml(customer.companyName || '-')}</td>
                <td>${this.escapeHtml(customer.phone || '-')}</td>
                <td>${this.escapeHtml(customer.email || '-')}</td>
                <td><span class="status-badge status-${customer.status}">${this.getStatusLabel(customer.status)}</span></td>
                <td><span class="rank-badge rank-${customer.rank}">${customer.rank}</span></td>
                <td>
                    <div class="table-actions">
                        <button class="btn btn-secondary btn-icon view-btn" title="詳細を見る">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path>
                                <circle cx="12" cy="12" r="3"></circle>
                            </svg>
                        </button>
                    </div>
                </td>
            </tr>
        `).join('');

        tbody.querySelectorAll('tr').forEach(row => {
            row.querySelector('.view-btn').addEventListener('click', (e) => {
                e.stopPropagation();
                this.openCustomerDetail(row.dataset.id);
            });
            row.addEventListener('click', () => {
                this.openCustomerDetail(row.dataset.id);
            });
        });
    }

    // ===== History List =====
    renderHistoryList(searchQuery = '') {
        const container = document.getElementById('historyList');
        const emptyState = document.getElementById('historyEmptyState');
        
        let filtered = [...this.history];
        
        // Search filter
        if (searchQuery) {
            const query = searchQuery.toLowerCase();
            filtered = filtered.filter(h => {
                const customer = this.customers.find(c => c.id === h.customerId);
                return h.subject.toLowerCase().includes(query) ||
                    h.content.toLowerCase().includes(query) ||
                    (customer && customer.name.toLowerCase().includes(query));
            });
        }
        
        // Type filter
        const typeFilter = document.getElementById('historyTypeFilter').value;
        if (typeFilter) {
            filtered = filtered.filter(h => h.type === typeFilter);
        }
        
        // Period filter
        const periodFilter = document.getElementById('historyPeriodFilter').value;
        if (periodFilter) {
            const now = new Date();
            let startDate;
            
            switch (periodFilter) {
                case 'today':
                    startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
                    break;
                case 'week':
                    startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
                    break;
                case 'month':
                    startDate = new Date(now.getFullYear(), now.getMonth(), 1);
                    break;
            }
            
            if (startDate) {
                filtered = filtered.filter(h => new Date(h.date) >= startDate);
            }
        }
        
        // Sort by date desc
        filtered.sort((a, b) => new Date(b.date) - new Date(a.date));
        
        if (filtered.length === 0) {
            container.innerHTML = '';
            emptyState.style.display = 'flex';
            return;
        }
        
        emptyState.style.display = 'none';
        
        container.innerHTML = filtered.map(h => {
            const customer = this.customers.find(c => c.id === h.customerId);
            return `
                <div class="history-item" data-id="${h.id}">
                    <div class="history-header">
                        <div class="history-meta">
                            <span class="history-type-badge ${h.type}">${this.getTypeIcon(h.type)} ${this.getTypeLabel(h.type)}</span>
                            <span class="history-customer" data-customer-id="${h.customerId}">${customer ? this.escapeHtml(customer.name) : '削除済み'}</span>
                        </div>
                        <div class="history-date">${this.formatDateTime(h.date)}</div>
                    </div>
                    <div class="history-subject">${this.escapeHtml(h.subject)}</div>
                    <div class="history-content">${this.escapeHtml(h.content).substring(0, 200)}${h.content.length > 200 ? '...' : ''}</div>
                </div>
            `;
        }).join('');

        container.querySelectorAll('.history-customer').forEach(el => {
            el.addEventListener('click', (e) => {
                e.stopPropagation();
                const customerId = el.dataset.customerId;
                if (customerId && this.customers.find(c => c.id === customerId)) {
                    this.openCustomerDetail(customerId);
                }
            });
        });
    }

    // ===== Customer Modal =====
    openCustomerModal(customer = null) {
        const modal = document.getElementById('customerModal');
        const title = document.getElementById('customerModalTitle');
        const form = document.getElementById('customerForm');
        
        form.reset();
        
        if (customer) {
            title.textContent = '顧客情報を編集';
            document.getElementById('customerId').value = customer.id;
            document.getElementById('customerName').value = customer.name;
            document.getElementById('customerNameKana').value = customer.nameKana || '';
            document.getElementById('companyName').value = customer.companyName || '';
            document.getElementById('department').value = customer.department || '';
            document.getElementById('position').value = customer.position || '';
            document.getElementById('phone').value = customer.phone || '';
            document.getElementById('mobile').value = customer.mobile || '';
            document.getElementById('email').value = customer.email || '';
            document.getElementById('address').value = customer.address || '';
            document.getElementById('status').value = customer.status || 'active';
            document.getElementById('rank').value = customer.rank || 'B';
            document.getElementById('notes').value = customer.notes || '';
        } else {
            title.textContent = '新規顧客登録';
            document.getElementById('customerId').value = '';
        }
        
        modal.classList.add('active');
    }

    closeCustomerModal() {
        document.getElementById('customerModal').classList.remove('active');
    }

    saveCustomer() {
        const id = document.getElementById('customerId').value;
        const customerData = {
            id: id || this.generateId(),
            name: document.getElementById('customerName').value.trim(),
            nameKana: document.getElementById('customerNameKana').value.trim(),
            companyName: document.getElementById('companyName').value.trim(),
            department: document.getElementById('department').value.trim(),
            position: document.getElementById('position').value.trim(),
            phone: document.getElementById('phone').value.trim(),
            mobile: document.getElementById('mobile').value.trim(),
            email: document.getElementById('email').value.trim(),
            address: document.getElementById('address').value.trim(),
            status: document.getElementById('status').value,
            rank: document.getElementById('rank').value,
            notes: document.getElementById('notes').value.trim(),
            createdAt: id ? this.customers.find(c => c.id === id)?.createdAt : new Date().toISOString(),
            updatedAt: new Date().toISOString()
        };

        if (id) {
            const index = this.customers.findIndex(c => c.id === id);
            if (index !== -1) {
                this.customers[index] = customerData;
            }
            this.showToast('顧客情報を更新しました', 'success');
        } else {
            this.customers.push(customerData);
            this.showToast('新規顧客を登録しました', 'success');
        }

        this.saveData();
        this.closeCustomerModal();
        this.renderDashboard();
        this.renderCustomersTable();
        this.updateStats();
        
        // If detail modal is open, refresh it
        if (this.currentCustomerId === customerData.id) {
            this.openCustomerDetail(customerData.id);
        }
    }

    // ===== Customer Detail Modal =====
    openCustomerDetail(customerId) {
        const customer = this.customers.find(c => c.id === customerId);
        if (!customer) return;

        this.currentCustomerId = customerId;
        const modal = document.getElementById('customerDetailModal');
        
        document.getElementById('customerDetailName').textContent = customer.name;
        
        // Render info
        const infoContainer = document.getElementById('customerDetailInfo');
        infoContainer.innerHTML = `
            <div class="detail-item">
                <div class="detail-label">顧客名</div>
                <div class="detail-value">${this.escapeHtml(customer.name)}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">フリガナ</div>
                <div class="detail-value">${this.escapeHtml(customer.nameKana || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">会社名</div>
                <div class="detail-value">${this.escapeHtml(customer.companyName || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">部署</div>
                <div class="detail-value">${this.escapeHtml(customer.department || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">役職</div>
                <div class="detail-value">${this.escapeHtml(customer.position || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">電話番号</div>
                <div class="detail-value">${this.escapeHtml(customer.phone || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">携帯電話</div>
                <div class="detail-value">${this.escapeHtml(customer.mobile || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">メール</div>
                <div class="detail-value">${customer.email ? `<a href="mailto:${this.escapeHtml(customer.email)}" style="color: var(--primary-light);">${this.escapeHtml(customer.email)}</a>` : '-'}</div>
            </div>
            <div class="detail-item full-width">
                <div class="detail-label">住所</div>
                <div class="detail-value">${this.escapeHtml(customer.address || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">ステータス</div>
                <div class="detail-value"><span class="status-badge status-${customer.status}">${this.getStatusLabel(customer.status)}</span></div>
            </div>
            <div class="detail-item">
                <div class="detail-label">ランク</div>
                <div class="detail-value"><span class="rank-badge rank-${customer.rank}">${customer.rank}</span></div>
            </div>
            <div class="detail-item full-width">
                <div class="detail-label">備考</div>
                <div class="detail-value">${this.escapeHtml(customer.notes || '-')}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">登録日</div>
                <div class="detail-value">${this.formatDateTime(customer.createdAt)}</div>
            </div>
            <div class="detail-item">
                <div class="detail-label">更新日</div>
                <div class="detail-value">${this.formatDateTime(customer.updatedAt)}</div>
            </div>
        `;

        // Render history for this customer
        this.renderCustomerHistory(customerId);
        
        // Reset tabs
        this.switchDetailTab('info');
        
        modal.classList.add('active');
    }

    closeDetailModal() {
        document.getElementById('customerDetailModal').classList.remove('active');
        this.currentCustomerId = null;
    }

    switchDetailTab(tabName) {
        document.querySelectorAll('.detail-tabs .tab-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabName);
        });
        document.getElementById('infoTab').classList.toggle('active', tabName === 'info');
        document.getElementById('customerHistoryTab').classList.toggle('active', tabName === 'customerHistory');
    }

    renderCustomerHistory(customerId) {
        const container = document.getElementById('customerHistoryList');
        const customerHistory = this.history
            .filter(h => h.customerId === customerId)
            .sort((a, b) => new Date(b.date) - new Date(a.date));
        
        if (customerHistory.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <p>この顧客の対応履歴はありません</p>
                </div>
            `;
            return;
        }

        container.innerHTML = customerHistory.map(h => `
            <div class="history-item" data-id="${h.id}">
                <div class="history-header">
                    <span class="history-type-badge ${h.type}">${this.getTypeIcon(h.type)} ${this.getTypeLabel(h.type)}</span>
                    <div class="history-actions">
                        <button class="btn btn-secondary btn-icon edit-history" title="編集">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                            </svg>
                        </button>
                        <button class="btn btn-danger btn-icon delete-history" title="削除">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <polyline points="3 6 5 6 21 6"></polyline>
                                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
                            </svg>
                        </button>
                    </div>
                </div>
                <div class="history-date" style="margin-bottom: 8px;">${this.formatDateTime(h.date)}</div>
                <div class="history-subject">${this.escapeHtml(h.subject)}</div>
                <div class="history-content">${this.escapeHtml(h.content)}</div>
                ${h.result ? `<div style="margin-top: 12px; padding-top: 12px; border-top: 1px solid var(--border-color);"><strong>結果・次回アクション:</strong><br>${this.escapeHtml(h.result)}</div>` : ''}
            </div>
        `).join('');

        container.querySelectorAll('.edit-history').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const historyId = btn.closest('.history-item').dataset.id;
                this.editHistory(historyId);
            });
        });

        container.querySelectorAll('.delete-history').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const historyId = btn.closest('.history-item').dataset.id;
                this.confirmDeleteHistory(historyId);
            });
        });
    }

    editCurrentCustomer() {
        const customer = this.customers.find(c => c.id === this.currentCustomerId);
        if (customer) {
            this.closeDetailModal();
            this.openCustomerModal(customer);
        }
    }

    confirmDeleteCustomer() {
        const customer = this.customers.find(c => c.id === this.currentCustomerId);
        if (!customer) return;

        this.showConfirmDialog(
            '顧客を削除',
            `「${customer.name}」を削除してもよろしいですか？この操作は取り消せません。`,
            () => {
                this.deleteCustomer(this.currentCustomerId);
            }
        );
    }

    deleteCustomer(customerId) {
        this.customers = this.customers.filter(c => c.id !== customerId);
        this.history = this.history.filter(h => h.customerId !== customerId);
        this.saveData();
        this.closeDetailModal();
        this.closeConfirmModal();
        this.renderDashboard();
        this.renderCustomersTable();
        this.renderHistoryList();
        this.updateStats();
        this.showToast('顧客を削除しました', 'success');
    }

    // ===== History Modal =====
    openHistoryModal(historyData = null) {
        const modal = document.getElementById('historyModal');
        const title = document.getElementById('historyModalTitle');
        const form = document.getElementById('historyForm');
        
        form.reset();
        
        if (historyData) {
            title.textContent = '対応履歴を編集';
            document.getElementById('historyId').value = historyData.id;
            document.getElementById('historyCustomerId').value = historyData.customerId;
            document.getElementById('historyDate').value = this.formatDateTimeLocal(historyData.date);
            document.getElementById('historyType').value = historyData.type;
            document.getElementById('historySubject').value = historyData.subject;
            document.getElementById('historyContent').value = historyData.content;
            document.getElementById('historyResult').value = historyData.result || '';
        } else {
            title.textContent = '対応履歴を追加';
            document.getElementById('historyId').value = '';
            document.getElementById('historyCustomerId').value = this.currentCustomerId;
            document.getElementById('historyDate').value = this.formatDateTimeLocal(new Date().toISOString());
        }
        
        modal.classList.add('active');
    }

    closeHistoryModal() {
        document.getElementById('historyModal').classList.remove('active');
    }

    saveHistory() {
        const id = document.getElementById('historyId').value;
        const historyData = {
            id: id || this.generateId(),
            customerId: document.getElementById('historyCustomerId').value,
            date: new Date(document.getElementById('historyDate').value).toISOString(),
            type: document.getElementById('historyType').value,
            subject: document.getElementById('historySubject').value.trim(),
            content: document.getElementById('historyContent').value.trim(),
            result: document.getElementById('historyResult').value.trim(),
            createdAt: id ? this.history.find(h => h.id === id)?.createdAt : new Date().toISOString()
        };

        if (id) {
            const index = this.history.findIndex(h => h.id === id);
            if (index !== -1) {
                this.history[index] = historyData;
            }
            this.showToast('対応履歴を更新しました', 'success');
        } else {
            this.history.push(historyData);
            this.showToast('対応履歴を追加しました', 'success');
        }

        this.saveData();
        this.closeHistoryModal();
        this.renderDashboard();
        this.renderHistoryList();
        this.updateStats();
        
        if (this.currentCustomerId) {
            this.renderCustomerHistory(this.currentCustomerId);
        }
    }

    editHistory(historyId) {
        const historyData = this.history.find(h => h.id === historyId);
        if (historyData) {
            this.openHistoryModal(historyData);
        }
    }

    confirmDeleteHistory(historyId) {
        this.showConfirmDialog(
            '履歴を削除',
            'この対応履歴を削除してもよろしいですか？',
            () => {
                this.deleteHistory(historyId);
            }
        );
    }

    deleteHistory(historyId) {
        this.history = this.history.filter(h => h.id !== historyId);
        this.saveData();
        this.closeConfirmModal();
        this.renderDashboard();
        this.renderHistoryList();
        this.updateStats();
        
        if (this.currentCustomerId) {
            this.renderCustomerHistory(this.currentCustomerId);
        }
        
        this.showToast('対応履歴を削除しました', 'success');
    }

    // ===== Confirm Dialog =====
    showConfirmDialog(title, message, onConfirm) {
        const modal = document.getElementById('confirmModal');
        document.getElementById('confirmTitle').textContent = title;
        document.getElementById('confirmMessage').textContent = message;
        
        const confirmBtn = document.getElementById('confirmOk');
        const newConfirmBtn = confirmBtn.cloneNode(true);
        confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);
        
        newConfirmBtn.addEventListener('click', onConfirm);
        
        modal.classList.add('active');
    }

    closeConfirmModal() {
        document.getElementById('confirmModal').classList.remove('active');
    }

    // ===== Toast Notifications =====
    showToast(message, type = 'info') {
        const container = document.getElementById('toastContainer');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        const icons = {
            success: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>',
            error: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>',
            info: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>'
        };
        
        toast.innerHTML = `
            <div class="toast-icon ${type}">${icons[type]}</div>
            <div class="toast-message">${this.escapeHtml(message)}</div>
        `;
        
        container.appendChild(toast);
        
        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(100%)';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }

    // ===== Utility Functions =====
    generateId() {
        return 'id_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    getInitials(name) {
        if (!name) return '?';
        return name.charAt(0).toUpperCase();
    }

    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    formatDate(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    }

    formatDateTime(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    }

    formatDateTimeLocal(dateString) {
        const date = new Date(dateString);
        return date.getFullYear() + '-' +
            String(date.getMonth() + 1).padStart(2, '0') + '-' +
            String(date.getDate()).padStart(2, '0') + 'T' +
            String(date.getHours()).padStart(2, '0') + ':' +
            String(date.getMinutes()).padStart(2, '0');
    }

    getStatusLabel(status) {
        const labels = {
            active: 'アクティブ',
            pending: '保留中',
            inactive: '非アクティブ'
        };
        return labels[status] || status;
    }

    getTypeLabel(type) {
        const labels = {
            phone: '電話',
            email: 'メール',
            visit: '訪問',
            meeting: '打ち合わせ',
            other: 'その他'
        };
        return labels[type] || type;
    }

    getTypeIcon(type) {
        const icons = {
            phone: '📞',
            email: '✉️',
            visit: '🚗',
            meeting: '🤝',
            other: '📝'
        };
        return icons[type] || '📋';
    }
}

// Initialize the app
document.addEventListener('DOMContentLoaded', () => {
    window.crmApp = new CRMApp();
});
