/**
 * @fileoverview Customer Relationship Management (CRM) Application
 *
 * A comprehensive CRM system for managing customer data and interaction history.
 * Features include customer CRUD operations, contact history tracking,
 * advanced filtering, search functionality, and dashboard analytics.
 *
 * Data is persisted in browser localStorage with keys:
 * - 'crm_customers': Array of customer objects
 * - 'crm_history': Array of history/interaction objects
 *
 * @author CRM Development Team
 * @version 1.0.0
 */

/**
 * @typedef {Object} Customer
 * @property {string} id - Unique identifier for the customer
 * @property {string} name - Customer's full name
 * @property {string} nameKana - Customer's name in katakana (フリガナ)
 * @property {string} companyName - Company name
 * @property {string} department - Department name
 * @property {string} position - Job position/title
 * @property {string} phone - Primary phone number
 * @property {string} mobile - Mobile phone number
 * @property {string} email - Email address
 * @property {string} address - Full address
 * @property {('active'|'pending'|'inactive')} status - Customer status
 * @property {('S'|'A'|'B'|'C')} rank - Customer rank/priority
 * @property {string} notes - Additional notes about the customer
 * @property {string} createdAt - ISO 8601 timestamp of creation
 * @property {string} updatedAt - ISO 8601 timestamp of last update
 */

/**
 * @typedef {Object} History
 * @property {string} id - Unique identifier for the history record
 * @property {string} customerId - ID of the associated customer
 * @property {string} date - ISO 8601 timestamp of the interaction
 * @property {('phone'|'email'|'visit'|'meeting'|'other')} type - Type of interaction
 * @property {string} subject - Brief subject/title of the interaction
 * @property {string} content - Detailed content of the interaction
 * @property {string} result - Results or next actions from the interaction
 * @property {string} createdAt - ISO 8601 timestamp of record creation
 */

/**
 * Main CRM Application Class
 *
 * Manages all customer relationship data, UI interactions, and business logic.
 * Implements a single-page application with multiple views (dashboard, customers, history).
 * All data operations are performed in-memory and persisted to localStorage.
 *
 * @class
 * @example
 * // Application is automatically initialized on DOMContentLoaded
 * document.addEventListener('DOMContentLoaded', () => {
 *     window.crmApp = new CRMApp();
 * });
 */
class CRMApp {
    /**
     * Initializes the CRM application with default state
     *
     * @constructor
     */
    constructor() {
        /**
         * Array of all customer records
         * @type {Customer[]}
         */
        this.customers = [];

        /**
         * Array of all interaction history records
         * @type {History[]}
         */
        this.history = [];

        /**
         * ID of the currently selected/viewed customer
         * @type {string|null}
         */
        this.currentCustomerId = null;

        this.init();
    }

    /**
     * Initializes the application by loading data and setting up the UI
     *
     * Called automatically by the constructor. Loads data from localStorage,
     * binds all event listeners, and renders all views with initial data.
     *
     * @private
     */
    init() {
        this.loadData();
        this.bindEvents();
        this.renderDashboard();
        this.renderCustomersTable();
        this.renderHistoryList();
        this.updateStats();
    }

    // ===== LocalStorage =====

    /**
     * Loads customer and history data from browser localStorage
     *
     * Attempts to load persisted data from localStorage keys 'crm_customers'
     * and 'crm_history'. If data exists, it's parsed and loaded into memory.
     * If no data exists, arrays remain empty.
     *
     * @private
     */
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

    /**
     * Persists current customer and history data to browser localStorage
     *
     * Serializes current in-memory data and saves to localStorage.
     * Called after any data modification operation (create, update, delete).
     *
     * @private
     */
    saveData() {
        localStorage.setItem('crm_customers', JSON.stringify(this.customers));
        localStorage.setItem('crm_history', JSON.stringify(this.history));
    }

    // ===== Event Bindings =====

    /**
     * Binds all event listeners for UI interactions
     *
     * Sets up event listeners for navigation, search, modal controls,
     * form submissions, and filter changes. Called once during initialization.
     *
     * @private
     */
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

    /**
     * Switches between different views (dashboard, customers, history)
     *
     * Updates navigation item and view active states based on the selected view.
     * Only one view is visible at a time.
     *
     * @param {string} viewName - Name of the view to display ('dashboard', 'customers', 'history')
     * @private
     */
    switchView(viewName) {
        document.querySelectorAll('.nav-item').forEach(item => {
            item.classList.toggle('active', item.dataset.view === viewName);
        });
        document.querySelectorAll('.view').forEach(view => {
            view.classList.toggle('active', view.id === viewName + 'View');
        });
    }

    // ===== Search =====

    /**
     * Handles global search functionality across different views
     *
     * Applies search query to the currently active view. Search is performed
     * on relevant fields depending on the view (customer names, emails, phone
     * numbers for customers view; subjects, content for history view).
     *
     * @param {string} query - Search query string (case-insensitive)
     * @private
     */
    handleSearch(query) {
        const currentView = document.querySelector('.view.active').id;

        if (currentView === 'customersView') {
            this.renderCustomersTable(query);
        } else if (currentView === 'historyView') {
            this.renderHistoryList(query);
        }
    }

    // ===== Dashboard =====

    /**
     * Updates dashboard statistics counters
     *
     * Calculates and displays various metrics including:
     * - Total customer count
     * - Active customer count
     * - Total history records
     * - Monthly history count (current month)
     * - Today's activity count
     *
     * @private
     */
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

    /**
     * Renders the complete dashboard view
     *
     * Triggers rendering of recent customers and recent history sections.
     *
     * @private
     */
    renderDashboard() {
        this.renderRecentCustomers();
        this.renderRecentHistory();
    }

    /**
     * Renders the list of 5 most recently created customers
     *
     * Displays customer cards sorted by creation date (newest first).
     * Shows customer name, company, and creation date. Clicking a card
     * opens the customer detail modal. Displays empty state if no customers exist.
     *
     * @private
     */
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

    /**
     * Renders the list of 5 most recent interaction history records
     *
     * Displays history cards sorted by interaction date (newest first).
     * Shows interaction subject, customer name, and timestamp. Clicking a card
     * opens the customer detail modal. Displays empty state if no history exists.
     *
     * @private
     */
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

    /**
     * Renders the customers table with filtering and sorting
     *
     * Displays all customers in a table format with support for:
     * - Text search across name, company, phone, and email fields
     * - Status filtering (active/pending/inactive)
     * - Rank filtering (S/A/B/C)
     * - Sorting by newest, oldest, name, or company
     *
     * Each row is clickable to open customer details.
     *
     * @param {string} [searchQuery=''] - Optional search query to filter customers
     * @private
     */
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

    /**
     * Renders the interaction history list with filtering
     *
     * Displays all interaction history records with support for:
     * - Text search across subject, content, and customer name
     * - Type filtering (phone/email/visit/meeting/other)
     * - Period filtering (today/week/month/all)
     * - Always sorted by date (newest first)
     *
     * History items show type badge, customer name, subject, and truncated content.
     * Customer names are clickable to open customer details.
     *
     * @param {string} [searchQuery=''] - Optional search query to filter history
     * @private
     */
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

    /**
     * Opens the customer form modal for creating or editing a customer
     *
     * If a customer object is provided, the form is populated for editing.
     * If no customer is provided, an empty form is shown for creating a new customer.
     *
     * @param {Customer|null} [customer=null] - Customer to edit, or null to create new
     * @public
     */
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

    /**
     * Closes the customer form modal
     *
     * @private
     */
    closeCustomerModal() {
        document.getElementById('customerModal').classList.remove('active');
    }

    /**
     * Saves customer data from the form (create or update)
     *
     * Reads form values, creates/updates the customer record, persists to localStorage,
     * and refreshes all relevant views. Shows a success toast notification.
     * Preserves createdAt timestamp for updates.
     *
     * @private
     */
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

    /**
     * Opens the customer detail modal showing full customer information
     *
     * Displays comprehensive customer information in a tabbed modal with:
     * - Info tab: All customer fields and metadata
     * - History tab: All interaction history for this customer
     *
     * @param {string} customerId - ID of the customer to display
     * @public
     */
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

    /**
     * Closes the customer detail modal
     *
     * Clears the currently selected customer ID.
     *
     * @private
     */
    closeDetailModal() {
        document.getElementById('customerDetailModal').classList.remove('active');
        this.currentCustomerId = null;
    }

    /**
     * Switches between tabs in the customer detail modal
     *
     * @param {string} tabName - Tab name to activate ('info' or 'customerHistory')
     * @private
     */
    switchDetailTab(tabName) {
        document.querySelectorAll('.detail-tabs .tab-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabName);
        });
        document.getElementById('infoTab').classList.toggle('active', tabName === 'info');
        document.getElementById('customerHistoryTab').classList.toggle('active', tabName === 'customerHistory');
    }

    /**
     * Renders interaction history for a specific customer
     *
     * Displays all history records associated with the customer, sorted by date (newest first).
     * Each history item includes edit and delete buttons. Shows empty state if no history exists.
     *
     * @param {string} customerId - ID of the customer whose history to display
     * @private
     */
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

    /**
     * Initiates editing of the currently viewed customer
     *
     * Closes the detail modal and opens the edit form modal with customer data.
     *
     * @private
     */
    editCurrentCustomer() {
        const customer = this.customers.find(c => c.id === this.currentCustomerId);
        if (customer) {
            this.closeDetailModal();
            this.openCustomerModal(customer);
        }
    }

    /**
     * Shows confirmation dialog for deleting the currently viewed customer
     *
     * @private
     */
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

    /**
     * Deletes a customer and all associated history records
     *
     * Removes the customer and all interaction history, persists changes,
     * closes modals, refreshes views, and shows success notification.
     *
     * @param {string} customerId - ID of the customer to delete
     * @private
     */
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

    /**
     * Opens the history form modal for creating or editing an interaction record
     *
     * If historyData is provided, the form is populated for editing.
     * If not provided, an empty form is shown with the current customer pre-selected.
     *
     * @param {History|null} [historyData=null] - History record to edit, or null to create new
     * @public
     */
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

    /**
     * Closes the history form modal
     *
     * @private
     */
    closeHistoryModal() {
        document.getElementById('historyModal').classList.remove('active');
    }

    /**
     * Saves history data from the form (create or update)
     *
     * Reads form values, creates/updates the history record, persists to localStorage,
     * and refreshes all relevant views. Shows a success toast notification.
     *
     * @private
     */
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

    /**
     * Initiates editing of a history record
     *
     * Opens the history form modal populated with the history data.
     *
     * @param {string} historyId - ID of the history record to edit
     * @private
     */
    editHistory(historyId) {
        const historyData = this.history.find(h => h.id === historyId);
        if (historyData) {
            this.openHistoryModal(historyData);
        }
    }

    /**
     * Shows confirmation dialog for deleting a history record
     *
     * @param {string} historyId - ID of the history record to delete
     * @private
     */
    confirmDeleteHistory(historyId) {
        this.showConfirmDialog(
            '履歴を削除',
            'この対応履歴を削除してもよろしいですか？',
            () => {
                this.deleteHistory(historyId);
            }
        );
    }

    /**
     * Deletes a history record
     *
     * Removes the history record, persists changes, closes modal,
     * refreshes views, and shows success notification.
     *
     * @param {string} historyId - ID of the history record to delete
     * @private
     */
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

    /**
     * Shows a confirmation dialog with custom title and message
     *
     * Displays a modal dialog with confirm and cancel buttons.
     * The onConfirm callback is executed when the user confirms.
     *
     * @param {string} title - Dialog title
     * @param {string} message - Confirmation message to display
     * @param {Function} onConfirm - Callback function to execute on confirmation
     * @private
     */
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

    /**
     * Closes the confirmation dialog
     *
     * @private
     */
    closeConfirmModal() {
        document.getElementById('confirmModal').classList.remove('active');
    }

    // ===== Toast Notifications =====

    /**
     * Displays a temporary toast notification
     *
     * Shows a notification message that auto-dismisses after 3 seconds.
     * Supports three types: success (green), error (red), and info (blue).
     *
     * @param {string} message - Message to display in the toast
     * @param {('success'|'error'|'info')} [type='info'] - Toast type/style
     * @private
     * @example
     * this.showToast('Customer saved successfully', 'success');
     * this.showToast('An error occurred', 'error');
     */
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

    /**
     * Generates a unique ID for new records
     *
     * Creates a unique identifier using current timestamp and random string.
     * Format: 'id_<timestamp>_<random>'
     *
     * @returns {string} Unique ID string
     * @private
     * @example
     * // Returns something like: 'id_1640000000000_a1b2c3d4e'
     * const id = this.generateId();
     */
    generateId() {
        return 'id_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    /**
     * Extracts initials from a name for avatar display
     *
     * Returns the first character of the name in uppercase.
     * Returns '?' if name is empty or undefined.
     *
     * @param {string} name - Full name to extract initials from
     * @returns {string} First character in uppercase, or '?'
     * @private
     * @example
     * this.getInitials('田中太郎'); // Returns '田'
     * this.getInitials('John Smith'); // Returns 'J'
     */
    getInitials(name) {
        if (!name) return '?';
        return name.charAt(0).toUpperCase();
    }

    /**
     * Escapes HTML special characters to prevent XSS attacks
     *
     * Converts HTML special characters to their entity equivalents.
     * Uses browser's built-in text node escaping for security.
     *
     * @param {string} text - Text to escape
     * @returns {string} HTML-escaped text
     * @private
     * @example
     * this.escapeHtml('<script>alert("xss")</script>');
     * // Returns '&lt;script&gt;alert("xss")&lt;/script&gt;'
     */
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * Formats an ISO date string to Japanese locale date format
     *
     * @param {string} dateString - ISO 8601 date string
     * @returns {string} Formatted date (e.g., '2024年1月15日')
     * @private
     * @example
     * this.formatDate('2024-01-15T10:30:00.000Z');
     * // Returns '2024年1月15日'
     */
    formatDate(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('ja-JP', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    }

    /**
     * Formats an ISO date string to Japanese locale datetime format
     *
     * @param {string} dateString - ISO 8601 date string
     * @returns {string} Formatted datetime (e.g., '2024年1月15日 10:30')
     * @private
     * @example
     * this.formatDateTime('2024-01-15T10:30:00.000Z');
     * // Returns '2024年1月15日 10:30'
     */
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

    /**
     * Formats an ISO date string to HTML datetime-local input format
     *
     * Converts ISO 8601 string to format required by HTML datetime-local inputs.
     * Format: 'YYYY-MM-DDTHH:mm'
     *
     * @param {string} dateString - ISO 8601 date string
     * @returns {string} Datetime in HTML5 datetime-local format
     * @private
     * @example
     * this.formatDateTimeLocal('2024-01-15T10:30:00.000Z');
     * // Returns '2024-01-15T10:30'
     */
    formatDateTimeLocal(dateString) {
        const date = new Date(dateString);
        return date.getFullYear() + '-' +
            String(date.getMonth() + 1).padStart(2, '0') + '-' +
            String(date.getDate()).padStart(2, '0') + 'T' +
            String(date.getHours()).padStart(2, '0') + ':' +
            String(date.getMinutes()).padStart(2, '0');
    }

    /**
     * Converts customer status code to Japanese label
     *
     * @param {string} status - Status code ('active', 'pending', 'inactive')
     * @returns {string} Japanese status label
     * @private
     */
    getStatusLabel(status) {
        const labels = {
            active: 'アクティブ',
            pending: '保留中',
            inactive: '非アクティブ'
        };
        return labels[status] || status;
    }

    /**
     * Converts interaction type code to Japanese label
     *
     * @param {string} type - Type code ('phone', 'email', 'visit', 'meeting', 'other')
     * @returns {string} Japanese type label
     * @private
     */
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

    /**
     * Gets emoji icon for interaction type
     *
     * @param {string} type - Type code ('phone', 'email', 'visit', 'meeting', 'other')
     * @returns {string} Emoji icon representing the type
     * @private
     */
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

/**
 * Application Initialization
 *
 * Automatically initializes the CRM application when the DOM is fully loaded.
 * Creates a global instance accessible via window.crmApp for debugging purposes.
 */
document.addEventListener('DOMContentLoaded', () => {
    window.crmApp = new CRMApp();
});
