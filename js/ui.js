/**
 * Módulo para gestión de interfaz de usuario y eventos
 */
const UIModule = {
    
    // Variables del módulo
    currentTab: 'parametros',
    notifications: [],
    
    /**
     * Inicializa la interfaz de usuario
     */
    init() {
        this.setupEventListeners();
        this.setupKeyboardShortcuts();
        this.setupTooltips();
        this.setupAutoSave();
        this.loadUIState();
    },

    /**
     * Configura event listeners para elementos de la UI
     */
    setupEventListeners() {
        // Event listeners para filtros con debounce
        const searchInputs = [
            'tableSearch',
            'repoSearch'
        ];
        
        searchInputs.forEach(inputId => {
            const input = document.getElementById(inputId);
            if (input) {
                input.addEventListener('input', 
                    Utils.debounce(() => this.handleSearch(inputId), 300)
                );
            }
        });

        // Event listener para cambios en CREATE TABLE input
        const createTableInput = document.getElementById('createTableInput');
        if (createTableInput) {
            createTableInput.addEventListener('paste', () => {
                setTimeout(() => this.handleCreateTablePaste(), 100);
            });
        }

        // Event listeners para validación en tiempo real
        this.setupRealTimeValidation();
        
        // Event listener para cerrar modal con Esc
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                this.closeModal();
            }
        });

        // Event listener para clicks fuera del modal
        const modal = document.getElementById('tableModal');
        if (modal) {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) {
                    this.closeModal();
                }
            });
        }
    },

    /**
     * Configura atajos de teclado
     */
    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            // Ctrl/Cmd + número para cambiar pestañas
            if ((e.ctrlKey || e.metaKey) && e.key >= '1' && e.key <= '5') {
                e.preventDefault();
                const tabNumber = parseInt(e.key);
                this.switchToTabByNumber(tabNumber);
            }
            
            // Ctrl/Cmd + S para guardar parámetros
            if ((e.ctrlKey || e.metaKey) && e.key === 's') {
                e.preventDefault();
                if (this.currentTab === 'parametros') {
                    ParametersModule.saveParameters();
                }
            }
            
            // Ctrl/Cmd + G para generar queries
            if ((e.ctrlKey || e.metaKey) && e.key === 'g') {
                e.preventDefault();
                QueryModule.generateAllQueries();
            }
            
            // Ctrl/Cmd + E para exportar
            if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
                e.preventDefault();
                ExportModule.exportToExcel();
            }
        });
    },

    /**
     * Cambia a una pestaña por número
     * @param {number} tabNumber - Número de pestaña (1-5)
     */
    switchToTabByNumber(tabNumber) {
        const tabMap = {
            1: 'parametros',
            2: 'describe',
            3: 'queries',
            4: 'repositorio',
            5: 'export'
        };
        
        const tabName = tabMap[tabNumber];
        if (tabName) {
            window.switchTab(tabName);
        }
    },

    /**
     * Configura tooltips informativos
     */
    setupTooltips() {
        // Agregar tooltips a elementos con atributo title
        const elementsWithTooltips = document.querySelectorAll('[title]');
        elementsWithTooltips.forEach(element => {
            this.enhanceTooltip(element);
        });
    },

    /**
     * Mejora tooltips básicos
     * @param {HTMLElement} element - Elemento con tooltip
     */
    enhanceTooltip(element) {
        let tooltip = null;
        
        element.addEventListener('mouseenter', (e) => {
            const title = e.target.getAttribute('title');
            if (!title) return;
            
            // Crear tooltip personalizado
            tooltip = document.createElement('div');
            tooltip.className = 'custom-tooltip';
            tooltip.textContent = title;
            document.body.appendChild(tooltip);
            
            // Posicionar tooltip
            this.positionTooltip(tooltip, e);
            
            // Ocultar title original
            e.target.setAttribute('data-original-title', title);
            e.target.removeAttribute('title');
        });
        
        element.addEventListener('mouseleave', (e) => {
            if (tooltip) {
                document.body.removeChild(tooltip);
                tooltip = null;
            }
            
            // Restaurar title original
            const originalTitle = e.target.getAttribute('data-original-title');
            if (originalTitle) {
                e.target.setAttribute('title', originalTitle);
                e.target.removeAttribute('data-original-title');
            }
        });
        
        element.addEventListener('mousemove', (e) => {
            if (tooltip) {
                this.positionTooltip(tooltip, e);
            }
        });
    },

    /**
     * Posiciona tooltip cerca del cursor
     * @param {HTMLElement} tooltip - Elemento tooltip
     * @param {Event} e - Evento del mouse
     */
    positionTooltip(tooltip, e) {
        const offset = 10;
        let x = e.pageX + offset;
        let y = e.pageY + offset;
        
        // Ajustar si se sale de la pantalla
        const tooltipRect = tooltip.getBoundingClientRect();
        if (x + tooltipRect.width > window.innerWidth) {
            x = e.pageX - tooltipRect.width - offset;
        }
        if (y + tooltipRect.height > window.innerHeight) {
            y = e.pageY - tooltipRect.height - offset;
        }
        
        tooltip.style.left = x + 'px';
        tooltip.style.top = y + 'px';
    },

    /**
     * Configura auto-guardado
     */
    setupAutoSave() {
        // Auto-guardar parámetros cada 30 segundos si hay cambios
        setInterval(() => {
            this.autoSaveParameters();
        }, 30000);
        
        // Guardar estado de UI antes de cerrar
        window.addEventListener('beforeunload', () => {
            this.saveUIState();
        });
    },

    /**
     * Auto-guarda parámetros si hay cambios pendientes
     */
    autoSaveParameters() {
        // Solo si estamos en la pestaña de parámetros y hay cambios
        if (this.currentTab === 'parametros' && this.hasUnsavedChanges()) {
            const params = this.collectParametersFromForm();
            localStorage.setItem('cuadreParameters_autosave', JSON.stringify(params));
            this.showNotification('💾 Parámetros auto-guardados', 'info', 2000);
        }
    },

    /**
     * Verifica si hay cambios no guardados
     * @returns {boolean} - true si hay cambios pendientes
     */
    hasUnsavedChanges() {
        const currentParams = ParametersModule.getCurrentParameters();
        const formParams = this.collectParametersFromForm();
        
        return JSON.stringify(currentParams) !== JSON.stringify(formParams);
    },

    /**
     * Recolecta parámetros del formulario
     * @returns {Object} - Parámetros del formulario
     */
    collectParametersFromForm() {
        return {
            esquemaDDV: document.getElementById('esquemaDDV')?.value.trim() || '',
            tablaDDV: document.getElementById('tablaDDV')?.value.trim() || '',
            esquemaEDV: document.getElementById('esquemaEDV')?.value.trim() || '',
            tablaEDV: document.getElementById('tablaEDV')?.value.trim() || '',
            periodos: document.getElementById('periodos')?.value.trim() || '',
            renameRules: Utils.parseRenameRules(document.getElementById('renameRules')?.value || '')
        };
    },

    /**
     * Configura validación en tiempo real
     */
    setupRealTimeValidation() {
        const inputsToValidate = [
            { id: 'esquemaDDV', validator: this.validateSchema },
            { id: 'esquemaEDV', validator: this.validateSchema },
            { id: 'tablaDDV', validator: this.validateTableName },
            { id: 'tablaEDV', validator: this.validateTableName },
            { id: 'periodos', validator: this.validatePeriods }
        ];
        
        inputsToValidate.forEach(({ id, validator }) => {
            const input = document.getElementById(id);
            if (input) {
                input.addEventListener('blur', () => {
                    this.validateField(input, validator);
                });
                
                input.addEventListener('input', 
                    Utils.debounce(() => this.validateField(input, validator), 500)
                );
            }
        });
    },

    /**
     * Valida un campo específico
     * @param {HTMLElement} input - Input a validar
     * @param {Function} validator - Función validadora
     */
    validateField(input, validator) {
        const value = input.value.trim();
        const isValid = validator(value);
        
        // Remover clases previas
        input.classList.remove('valid', 'invalid');
        
        // Agregar clase según validación
        if (value && isValid) {
            input.classList.add('valid');
        } else if (value && !isValid) {
            input.classList.add('invalid');
        }
    },

    /**
     * Validadores específicos
     */
    validateSchema: (value) => value.includes('.') && !value.includes(' '),
    validateTableName: (value) => /^[a-zA-Z_][a-zA-Z0-9_]*$/.test(value),
    validatePeriods: (value) => /^\d{6}(\s*,\s*\d{6})*$/.test(value),

    /**
     * Maneja eventos de búsqueda
     * @param {string} inputId - ID del input de búsqueda
     */
    handleSearch(inputId) {
        switch (inputId) {
            case 'tableSearch':
                RepositoryModule.filterTables();
                break;
            case 'repoSearch':
                RepositoryModule.filterRepository();
                break;
        }
    },

    /**
     * Maneja paste en CREATE TABLE input
     */
    handleCreateTablePaste() {
        const input = document.getElementById('createTableInput');
        const value = input.value.trim();
        
        if (value && Utils.isValidCreateTable(value)) {
            // Sugerir auto-detección
            this.showNotification(
                '💡 CREATE TABLE detectado. ¿Quieres analizarlo automáticamente?',
                'info',
                5000,
                [
                    {
                        text: 'Sí, analizar',
                        action: () => TableAnalysisModule.detectCreateTable()
                    },
                    {
                        text: 'No, gracias',
                        action: () => this.dismissNotification()
                    }
                ]
            );
        }
    },

    /**
     * Guarda estado de la UI
     */
    saveUIState() {
        const uiState = {
            currentTab: this.currentTab,
            timestamp: new Date().toISOString(),
            filters: {
                schemaFilter: document.getElementById('schemaFilter')?.value || '',
                tableFilter: document.getElementById('tableFilter')?.value || '',
                showFilter: document.getElementById('showFilter')?.value || 'all'
            }
        };
        
        localStorage.setItem('cuadreUIState', JSON.stringify(uiState));
    },

    /**
     * Carga estado de la UI
     */
    loadUIState() {
        try {
            const stored = localStorage.getItem('cuadreUIState');
            if (stored) {
                const uiState = JSON.parse(stored);
                
                // Restaurar filtros
                if (uiState.filters) {
                    Object.entries(uiState.filters).forEach(([filterId, value]) => {
                        const element = document.getElementById(filterId);
                        if (element && value) {
                            element.value = value;
                        }
                    });
                }
                
                // Restaurar pestaña activa si fue reciente (menos de 1 hora)
                const lastTimestamp = new Date(uiState.timestamp);
                const now = new Date();
                const hoursDiff = (now - lastTimestamp) / (1000 * 60 * 60);
                
                if (hoursDiff < 1 && uiState.currentTab) {
                    setTimeout(() => {
                        window.switchTab(uiState.currentTab);
                    }, 100);
                }
            }
        } catch (error) {
            console.error('Error cargando estado de UI:', error);
        }
    },

    /**
     * Actualiza la pestaña actual
     * @param {string} tabName - Nombre de la pestaña
     */
    updateCurrentTab(tabName) {
        this.currentTab = tabName;
        this.saveUIState();
    },

    /**
     * Muestra notificación en la interfaz
     * @param {string} message - Mensaje a mostrar
     * @param {string} type - Tipo de notificación ('info', 'success', 'warning', 'error')
     * @param {number} duration - Duración en ms (0 = permanente)
     * @param {Array} actions - Array de acciones {text, action}
     */
    showNotification(message, type = 'info', duration = 5000, actions = []) {
        const notification = {
            id: Date.now() + Math.random(),
            message,
            type,
            timestamp: new Date(),
            actions
        };
        
        this.notifications.push(notification);
        this.renderNotification(notification);
        
        if (duration > 0) {
            setTimeout(() => {
                this.removeNotification(notification.id);
            }, duration);
        }
    },

    /**
     * Renderiza una notificación
     * @param {Object} notification - Objeto de notificación
     */
    renderNotification(notification) {
        let container = document.getElementById('notifications-container');
        
        if (!container) {
            container = document.createElement('div');
            container.id = 'notifications-container';
            container.className = 'notifications-container';
            document.body.appendChild(container);
        }
        
        const notificationEl = document.createElement('div');
        notificationEl.className = `notification notification-${notification.type}`;
        notificationEl.id = `notification-${notification.id}`;
        
        let actionsHTML = '';
        if (notification.actions && notification.actions.length > 0) {
            actionsHTML = '<div class="notification-actions">';
            notification.actions.forEach((action, index) => {
                actionsHTML += `<button class="notification-btn" onclick="UIModule.executeNotificationAction(${notification.id}, ${index})">${action.text}</button>`;
            });
            actionsHTML += '</div>';
        }
        
        notificationEl.innerHTML = `
            <div class="notification-content">
                <div class="notification-message">${notification.message}</div>
                ${actionsHTML}
            </div>
            <button class="notification-close" onclick="UIModule.removeNotification(${notification.id})">&times;</button>
        `;
        
        container.appendChild(notificationEl);
        
        // Animación de entrada
        setTimeout(() => {
            notificationEl.classList.add('notification-show');
        }, 10);
    },

    /**
     * Ejecuta acción de notificación
     * @param {number} notificationId - ID de la notificación
     * @param {number} actionIndex - Índice de la acción
     */
    executeNotificationAction(notificationId, actionIndex) {
        const notification = this.notifications.find(n => n.id === notificationId);
        if (notification && notification.actions[actionIndex]) {
            notification.actions[actionIndex].action();
            this.removeNotification(notificationId);
        }
    },

    /**
     * Remueve una notificación
     * @param {number} notificationId - ID de la notificación
     */
    removeNotification(notificationId) {
        const notificationEl = document.getElementById(`notification-${notificationId}`);
        if (notificationEl) {
            notificationEl.classList.add('notification-hide');
            setTimeout(() => {
                if (notificationEl.parentNode) {
                    notificationEl.parentNode.removeChild(notificationEl);
                }
            }, 300);
        }
        
        // Remover de la lista
        this.notifications = this.notifications.filter(n => n.id !== notificationId);
    },

    /**
     * Dismiss notificación actual
     */
    dismissNotification() {
        if (this.notifications.length > 0) {
            const lastNotification = this.notifications[this.notifications.length - 1];
            this.removeNotification(lastNotification.id);
        }
    },

    /**
     * Cierra modal
     */
    closeModal() {
        const modal = document.getElementById('tableModal');
        if (modal) {
            modal.style.display = 'none';
        }
    },

    /**
     * Muestra modal con contenido
     * @param {string} title - Título del modal
     * @param {string} content - Contenido del modal
     */
    showModal(title, content) {
        document.getElementById('modalTitle').textContent = title;
        document.getElementById('modalContent').innerHTML = content;
        document.getElementById('tableModal').style.display = 'block';
    },

    /**
     * Muestra loading spinner
     * @param {string} message - Mensaje de carga
     */
    showLoading(message = 'Cargando...') {
        let loader = document.getElementById('global-loader');
        
        if (!loader) {
            loader = document.createElement('div');
            loader.id = 'global-loader';
            loader.className = 'global-loader';
            loader.innerHTML = `
                <div class="loader-content">
                    <div class="loader-spinner"></div>
                    <div class="loader-message">${message}</div>
                </div>
            `;
            document.body.appendChild(loader);
        } else {
            loader.querySelector('.loader-message').textContent = message;
        }
        
        loader.style.display = 'flex';
    },

    /**
     * Oculta loading spinner
     */
    hideLoading() {
        const loader = document.getElementById('global-loader');
        if (loader) {
            loader.style.display = 'none';
        }
    },

    /**
     * Confirma acción destructiva
     * @param {string} message - Mensaje de confirmación
     * @param {Function} onConfirm - Función a ejecutar si se confirma
     * @param {Function} onCancel - Función a ejecutar si se cancela (opcional)
     */
    confirmDestructiveAction(message, onConfirm, onCancel = null) {
        const confirmed = confirm(`⚠️ ${message}\n\nEsta acción no se puede deshacer.`);
        
        if (confirmed) {
            onConfirm();
        } else if (onCancel) {
            onCancel();
        }
    },

    /**
     * Maneja errores globales
     * @param {Error} error - Error a manejar
     * @param {string} context - Contexto donde ocurrió el error
     */
    handleError(error, context = 'Operación') {
        console.error(`Error en ${context}:`, error);
        
        this.showNotification(
            `❌ Error en ${context}: ${error.message}`,
            'error',
            10000,
            [
                {
                    text: 'Ver detalles',
                    action: () => this.showModal('Error Details', `
                        <strong>Contexto:</strong> ${context}<br>
                        <strong>Mensaje:</strong> ${error.message}<br>
                        <strong>Stack:</strong><br>
                        <pre>${error.stack || 'No disponible'}</pre>
                    `)
                }
            ]
        );
    },

    /**
     * Muestra progreso de operación
     * @param {number} progress - Progreso (0-100)
     * @param {string} message - Mensaje de progreso
     */
    showProgress(progress, message = '') {
        let progressBar = document.getElementById('global-progress');
        
        if (!progressBar) {
            progressBar = document.createElement('div');
            progressBar.id = 'global-progress';
            progressBar.className = 'global-progress';
            progressBar.innerHTML = `
                <div class="progress-content">
                    <div class="progress-message">${message}</div>
                    <div class="progress-bar">
                        <div class="progress-fill"></div>
                    </div>
                    <div class="progress-text">0%</div>
                </div>
            `;
            document.body.appendChild(progressBar);
        }
        
        const fill = progressBar.querySelector('.progress-fill');
        const text = progressBar.querySelector('.progress-text');
        const messageEl = progressBar.querySelector('.progress-message');
        
        fill.style.width = `${Math.min(100, Math.max(0, progress))}%`;
        text.textContent = `${Math.round(progress)}%`;
        messageEl.textContent = message;
        
        progressBar.style.display = progress < 100 ? 'flex' : 'none';
    }
};

// Hacer funciones disponibles globalmente para event handlers
window.UIModule = UIModule;