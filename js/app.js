/**
 * Módulo principal de la aplicación
 * Coordina todos los demás módulos y maneja la inicialización
 */
const App = {
    
    // Variables globales de la aplicación
    version: '2.0.0',
    isInitialized: false,
    modules: {},
    
    /**
     * Inicializa la aplicación
     */
    async init() {
        try {
            UIModule.showLoading('Inicializando aplicación...');
            
            // Registrar módulos
            this.registerModules();
            
            // Inicializar módulos en orden
            await this.initializeModules();
            
            // Cargar datos iniciales
            this.loadInitialData();
            
            // Configurar manejo global de errores
            this.setupGlobalErrorHandling();
            
            // Marcar como inicializado
            this.isInitialized = true;
            
            UIModule.hideLoading();
            UIModule.showNotification(
                `🚀 Aplicación inicializada correctamente (v${this.version})`,
                'success',
                3000
            );
            
            console.log(`🏗️ Generador de Queries de Ratificación v${this.version} - Inicializado`);
            
        } catch (error) {
            UIModule.hideLoading();
            UIModule.handleError(error, 'Inicialización de aplicación');
        }
    },

    /**
     * Registra todos los módulos disponibles
     */
    registerModules() {
        this.modules = {
            ui: UIModule,
            parameters: ParametersModule,
            repository: RepositoryModule,
            tableAnalysis: TableAnalysisModule,
            query: QueryModule,
            export: ExportModule
        };
        
        console.log('📦 Módulos registrados:', Object.keys(this.modules));
    },

    /**
     * Inicializa módulos en el orden correcto
     */
    async initializeModules() {
        const initOrder = [
            'ui',
            'parameters', 
            'repository',
            'tableAnalysis',
            'query',
            'export'
        ];
        
        for (const moduleName of initOrder) {
            const module = this.modules[moduleName];
            if (module && typeof module.init === 'function') {
                UIModule.showLoading(`Inicializando ${moduleName}...`);
                await module.init();
                console.log(`✅ Módulo ${moduleName} inicializado`);
            }
        }
    },

    /**
     * Carga datos iniciales de la aplicación
     */
    loadInitialData() {
        // Cargar parámetros guardados
        ParametersModule.loadParametersFromStorage();
        
        // Cargar datos de ejemplo si no hay parámetros
        const currentParams = ParametersModule.getCurrentParameters();
        if (!currentParams || !currentParams.esquemaDDV) {
            ParametersModule.loadExampleData();
        }
        
        // Verificar si hay auto-guardado
        this.checkAutoSave();
        
        console.log('📊 Datos iniciales cargados');
    },

    /**
     * Verifica si hay datos auto-guardados
     */
    checkAutoSave() {
        try {
            const autoSaved = localStorage.getItem('cuadreParameters_autosave');
            if (autoSaved) {
                const autoSavedParams = JSON.parse(autoSaved);
                const currentParams = ParametersModule.getCurrentParameters();
                
                // Si los datos auto-guardados son diferentes a los actuales
                if (JSON.stringify(autoSavedParams) !== JSON.stringify(currentParams)) {
                    UIModule.showNotification(
                        '💾 Se encontraron datos auto-guardados. ¿Quieres restaurarlos?',
                        'info',
                        0, // Permanente hasta que el usuario decida
                        [
                            {
                                text: 'Restaurar',
                                action: () => this.restoreAutoSave(autoSavedParams)
                            },
                            {
                                text: 'Descartar',
                                action: () => this.discardAutoSave()
                            }
                        ]
                    );
                }
            }
        } catch (error) {
            console.error('Error verificando auto-guardado:', error);
        }
    },

    /**
     * Restaura datos auto-guardados
     * @param {Object} autoSavedParams - Parámetros auto-guardados
     */
    restoreAutoSave(autoSavedParams) {
        ParametersModule.parameters = autoSavedParams;
        ParametersModule.populateForm();
        localStorage.removeItem('cuadreParameters_autosave');
        UIModule.showNotification('✅ Datos auto-guardados restaurados', 'success', 3000);
    },

    /**
     * Descarta datos auto-guardados
     */
    discardAutoSave() {
        localStorage.removeItem('cuadreParameters_autosave');
        UIModule.showNotification('🗑️ Datos auto-guardados descartados', 'info', 3000);
    },

    /**
     * Configura manejo global de errores
     */
    setupGlobalErrorHandling() {
        // Errores JavaScript no manejados
        window.addEventListener('error', (event) => {
            console.error('Error JavaScript no manejado:', event.error);
            UIModule.handleError(event.error, 'JavaScript');
        });
        
        // Promesas rechazadas no manejadas
        window.addEventListener('unhandledrejection', (event) => {
            console.error('Promesa rechazada no manejada:', event.reason);
            UIModule.handleError(new Error(event.reason), 'Promise');
        });
        
        // Wrapper para funciones críticas
        this.wrapCriticalFunctions();
    },

    /**
     * Envuelve funciones críticas con manejo de errores
     */
    wrapCriticalFunctions() {
        const criticalFunctions = [
            { module: ParametersModule, method: 'saveParameters' },
            { module: TableAnalysisModule, method: 'parseCreateTable' },
            { module: QueryModule, method: 'generateAllQueries' },
            { module: ExportModule, method: 'exportToExcel' },
            { module: RepositoryModule, method: 'saveToRepository' }
        ];
        
        criticalFunctions.forEach(({ module, method }) => {
            if (module[method]) {
                const originalMethod = module[method];
                module[method] = function(...args) {
                    try {
                        return originalMethod.apply(this, args);
                    } catch (error) {
                        UIModule.handleError(error, `${module.constructor.name}.${method}`);
                        throw error;
                    }
                };
            }
        });
    },

    /**
     * Ejecuta flujo completo de cuadre
     */
    async executeCompleteFlow() {
        try {
            UIModule.showLoading('Ejecutando flujo completo...');
            
            // Paso 1: Validar parámetros
            UIModule.showProgress(20, 'Validando parámetros...');
            if (!ParametersModule.areParametersReady()) {
                throw new Error('Los parámetros no están configurados correctamente');
            }
            
            // Paso 2: Validar estructura de tabla
            UIModule.showProgress(40, 'Validando estructura de tabla...');
            const tableValidation = TableAnalysisModule.validateTableStructure();
            if (!tableValidation.isValid) {
                throw new Error('La estructura de tabla no es válida: ' + tableValidation.errors.join(', '));
            }
            
            // Paso 3: Generar queries
            UIModule.showProgress(60, 'Generando queries...');
            QueryModule.generateAllQueries();
            
            // Paso 4: Exportar a Excel
            UIModule.showProgress(80, 'Exportando a Excel...');
            ExportModule.exportToExcel();
            
            // Paso 5: Completado
            UIModule.showProgress(100, 'Completado');
            
            setTimeout(() => {
                UIModule.hideLoading();
                UIModule.showNotification(
                    '🎉 Flujo completo ejecutado exitosamente',
                    'success',
                    5000
                );
            }, 1000);
            
        } catch (error) {
            UIModule.hideLoading();
            UIModule.handleError(error, 'Flujo completo');
        }
    },

    /**
     * Ejecuta diagnóstico del sistema
     */
    runDiagnostics() {
        const diagnostics = {
            timestamp: new Date().toISOString(),
            version: this.version,
            browser: this.getBrowserInfo(),
            modules: this.checkModulesHealth(),
            data: this.checkDataIntegrity(),
            storage: this.checkStorageHealth(),
            performance: this.checkPerformance()
        };
        
        console.log('🔧 Diagnóstico del sistema:', diagnostics);
        
        UIModule.showModal('Diagnóstico del Sistema', `
            <div class="diagnostics-report">
                <h4>Información General</h4>
                <p><strong>Versión:</strong> ${diagnostics.version}</p>
                <p><strong>Navegador:</strong> ${diagnostics.browser.name} ${diagnostics.browser.version}</p>
                <p><strong>Fecha:</strong> ${new Date(diagnostics.timestamp).toLocaleString()}</p>
                
                <h4>Estado de Módulos</h4>
                ${Object.entries(diagnostics.modules).map(([module, status]) => 
                    `<p><strong>${module}:</strong> ${status ? '✅' : '❌'}</p>`
                ).join('')}
                
                <h4>Integridad de Datos</h4>
                <p><strong>Parámetros:</strong> ${diagnostics.data.parameters ? '✅' : '❌'}</p>
                <p><strong>Repositorio:</strong> ${diagnostics.data.repository ? '✅' : '❌'}</p>
                <p><strong>Estructura de tabla:</strong> ${diagnostics.data.tableStructure ? '✅' : '❌'}</p>
                
                <h4>Almacenamiento</h4>
                <p><strong>LocalStorage disponible:</strong> ${diagnostics.storage.available ? '✅' : '❌'}</p>
                <p><strong>Datos guardados:</strong> ${diagnostics.storage.dataSize} KB</p>
                
                <h4>Rendimiento</h4>
                <p><strong>Tiempo de carga:</strong> ${diagnostics.performance.loadTime}ms</p>
                <p><strong>Memoria utilizada:</strong> ${diagnostics.performance.memoryUsage}</p>
            </div>
        `);
        
        return diagnostics;
    },

    /**
     * Obtiene información del navegador
     * @returns {Object} - Información del navegador
     */
    getBrowserInfo() {
        const ua = navigator.userAgent;
        let browserName = 'Unknown';
        let browserVersion = 'Unknown';
        
        if (ua.includes('Chrome')) {
            browserName = 'Chrome';
            browserVersion = ua.match(/Chrome\/([0-9.]+)/)?.[1] || 'Unknown';
        } else if (ua.includes('Firefox')) {
            browserName = 'Firefox';
            browserVersion = ua.match(/Firefox\/([0-9.]+)/)?.[1] || 'Unknown';
        } else if (ua.includes('Safari')) {
            browserName = 'Safari';
            browserVersion = ua.match(/Version\/([0-9.]+)/)?.[1] || 'Unknown';
        } else if (ua.includes('Edge')) {
            browserName = 'Edge';
            browserVersion = ua.match(/Edge\/([0-9.]+)/)?.[1] || 'Unknown';
        }
        
        return { name: browserName, version: browserVersion };
    },

    /**
     * Verifica salud de módulos
     * @returns {Object} - Estado de cada módulo
     */
    checkModulesHealth() {
        const health = {};
        
        Object.entries(this.modules).forEach(([name, module]) => {
            health[name] = module && typeof module === 'object';
        });
        
        return health;
    },

    /**
     * Verifica integridad de datos
     * @returns {Object} - Estado de integridad de datos
     */
    checkDataIntegrity() {
        const params = ParametersModule.getCurrentParameters();
        const repository = RepositoryModule.getRepository();
        const tableStructure = TableAnalysisModule.getTableStructure();
        
        return {
            parameters: params && Object.keys(params).length > 0,
            repository: repository && Object.keys(repository).length >= 0,
            tableStructure: Array.isArray(tableStructure)
        };
    },

    /**
     * Verifica salud del almacenamiento
     * @returns {Object} - Estado del almacenamiento
     */
    checkStorageHealth() {
        let available = false;
        let dataSize = 0;
        
        try {
            localStorage.setItem('test', 'test');
            localStorage.removeItem('test');
            available = true;
            
            // Calcular tamaño aproximado de datos guardados
            const data = [
                localStorage.getItem('cuadreParameters'),
                localStorage.getItem('tablesRepository'),
                localStorage.getItem('cuadreUIState')
            ].filter(Boolean);
            
            dataSize = data.reduce((total, item) => total + item.length, 0) / 1024; // KB
            
        } catch (error) {
            available = false;
        }
        
        return {
            available,
            dataSize: Math.round(dataSize * 100) / 100
        };
    },

    /**
     * Verifica rendimiento básico
     * @returns {Object} - Métricas de rendimiento
     */
    checkPerformance() {
        const loadTime = performance.now();
        let memoryUsage = 'No disponible';
        
        if (performance.memory) {
            const memory = performance.memory;
            memoryUsage = `${Math.round(memory.usedJSHeapSize / 1024 / 1024)} MB`;
        }
        
        return {
            loadTime: Math.round(loadTime),
            memoryUsage
        };
    },

    /**
     * Resetea la aplicación a estado inicial
     */
    resetApplication() {
        UIModule.confirmDestructiveAction(
            'Esto eliminará todos los datos guardados y reiniciará la aplicación',
            () => {
                try {
                    // Limpiar localStorage
                    const keysToRemove = [
                        'cuadreParameters',
                        'cuadreParameters_autosave',
                        'tablesRepository',
                        'cuadreUIState'
                    ];
                    
                    keysToRemove.forEach(key => {
                        localStorage.removeItem(key);
                    });
                    
                    // Resetear módulos
                    Object.values(this.modules).forEach(module => {
                        if (typeof module.reset === 'function') {
                            module.reset();
                        }
                    });
                    
                    // Recargar página
                    setTimeout(() => {
                        window.location.reload();
                    }, 1000);
                    
                    UIModule.showNotification('🔄 Aplicación reiniciada', 'success', 3000);
                    
                } catch (error) {
                    UIModule.handleError(error, 'Reset de aplicación');
                }
            }
        );
    },

    /**
     * Exporta configuración completa de la aplicación
     */
    exportConfiguration() {
        try {
            const config = {
                version: this.version,
                timestamp: new Date().toISOString(),
                parameters: ParametersModule.getCurrentParameters(),
                repository: RepositoryModule.getRepository(),
                tableStructure: TableAnalysisModule.getTableStructure(),
                uiState: JSON.parse(localStorage.getItem('cuadreUIState') || '{}')
            };
            
            const filename = `cuadre_config_${new Date().toISOString().split('T')[0]}.json`;
            const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            
            UIModule.showNotification('📤 Configuración exportada', 'success', 3000);
            
        } catch (error) {
            UIModule.handleError(error, 'Export de configuración');
        }
    },

    /**
     * Importa configuración de la aplicación
     * @param {File} file - Archivo de configuración
     */
    importConfiguration(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const config = JSON.parse(e.target.result);
                
                // Validar estructura
                if (!config.version || !config.parameters) {
                    throw new Error('Archivo de configuración inválido');
                }
                
                // Confirmar importación
                UIModule.confirmDestructiveAction(
                    'Esto sobrescribirá la configuración actual',
                    () => {
                        // Importar datos
                        if (config.parameters) {
                            ParametersModule.importParameters(JSON.stringify(config.parameters));
                        }
                        
                        if (config.repository) {
                            RepositoryModule.tablesRepository = config.repository;
                            RepositoryModule.saveRepositoryToStorage();
                            RepositoryModule.updateUI();
                        }
                        
                        if (config.tableStructure) {
                            TableAnalysisModule.tableStructure = config.tableStructure;
                            TableAnalysisModule.displayFieldMapping();
                        }
                        
                        UIModule.showNotification('📥 Configuración importada correctamente', 'success', 5000);
                    }
                );
                
            } catch (error) {
                UIModule.handleError(error, 'Import de configuración');
            }
        };
        
        reader.readAsText(file);
    },

    /**
     * Obtiene información de estado de la aplicación
     * @returns {Object} - Estado actual de la aplicación
     */
    getAppState() {
        return {
            version: this.version,
            initialized: this.isInitialized,
            modules: Object.keys(this.modules),
            timestamp: new Date().toISOString()
        };
    }
};

// Inicializar aplicación cuando el DOM esté listo
document.addEventListener('DOMContentLoaded', () => {
    App.init();
});

// Funciones globales para compatibilidad con HTML
window.App = App;

// Sobrescribir la función switchTab original para integrar con UIModule
window.switchTab = function(tabName) {
    Utils.switchTab(tabName, { target: document.querySelector(`.tab:nth-child(${getTabIndex(tabName)})`) });
    UIModule.updateCurrentTab(tabName);
};

// Helper para obtener índice de pestaña
function getTabIndex(tabName) {
    const tabMap = {
        'parametros': 1,
        'describe': 2,
        'queries': 3,
        'repositorio': 4,
        'export': 5
    };
    return tabMap[tabName] || 1;
}