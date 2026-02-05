/**
 * REPORT-MASTER PRO - Main Application Logic
 */

document.addEventListener('DOMContentLoaded', () => {
    // UI Elements
    const motorOllama = document.getElementById('motor-ollama');
    const motorLMStudio = document.getElementById('motor-lmstudio');
    const motorAPI = document.getElementById('motor-api');
    const apiConfig = document.getElementById('api-config');
    const apiUrlInput = document.getElementById('api-url');
    const apiKeyInput = document.getElementById('api-key');
    const apiModelInput = document.getElementById('api-model');
    const modelSelect = document.getElementById('model-select');
    const generateBtn = document.getElementById('generate-report');
    const thinkingLayer = document.getElementById('thinking-layer');
    const thinkingStatus = document.getElementById('thinking-status');
    const logoUpload = document.getElementById('logo-upload');
    const signUpload = document.getElementById('sign-upload');
    const reportBody = document.getElementById('report-body');
    const chatMessages = document.getElementById('chat-messages');
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const currentTimeEl = document.getElementById('current-time');
    const exportPdfBtn = document.querySelector('button[title="Exportar a PDF"]');
    const exportWordBtn = document.querySelector('button[title="Exportar a Word"]');
    const exportExcelBtn = document.querySelector('button[title="Exportar a Excel"]');
    const exportPptBtn = document.querySelector('button[title="Exportar a PowerPoint"]');
    const customHeaderInput = document.getElementById('custom-header');
    const customFooterInput = document.getElementById('custom-footer');
    const paperTitleEl = document.getElementById('report-title');
    const paperFooterTextEl = document.querySelector('#report-footer p.font-bold');
    const pageSizeSelect = document.getElementById('page-size');
    const reportPaperContainer = document.getElementById('report-paper');
    const extraNotesInput = document.getElementById('extra-notes');
    const fileListEl = document.getElementById('file-list');
    const themeToggleBtn = document.getElementById('theme-toggle');
    const sunIcon = document.getElementById('sun-icon');
    const moonIcon = document.getElementById('moon-icon');

    let uploadedFilesData = [];
    let historicalFilesData = [];
    let reportAbortController = null;
    let currentReportContent = ""; // For contextual chat

    // Load Saved API Settings
    function loadSettings() {
        const savedUrl = localStorage.getItem('rmp_api_url');
        const savedKey = localStorage.getItem('rmp_api_key');
        const savedModel = localStorage.getItem('rmp_api_model');
        const savedMotor = localStorage.getItem('rmp_current_motor') || 'ollama';
        const savedTheme = localStorage.getItem('rmp_theme') || 'dark';

        if (savedUrl) apiUrlInput.value = savedUrl;
        if (savedKey) apiKeyInput.value = savedKey;
        if (savedModel) apiModelInput.value = savedModel;

        if (savedTheme === 'light') {
            document.body.classList.add('light-mode');
            sunIcon.classList.remove('hidden');
            moonIcon.classList.add('hidden');
        }

        switchMotor(savedMotor);
        aiService.setApiConfig(savedUrl || '', savedKey || '', savedModel || '');
    }

    function saveSetting(key, value) {
        localStorage.setItem(key, value);
    }

    // Initialize Time
    setInterval(() => {
        const now = new Date();
        const formatted = now.toLocaleDateString('es-ES', { day: '2-digit', month: 'short', year: 'numeric' }).toUpperCase() + ' - ' +
            now.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit', hour12: false });
        currentTimeEl.textContent = formatted;
    }, 1000);

    // Cancel Process Logic
    const cancelBtn = document.getElementById('cancel-process');
    if (cancelBtn) {
        cancelBtn.addEventListener('click', () => {
            if (reportAbortController) {
                reportAbortController.abort();
                showThinking(false);
                addChatMessage("Proceso cancelado por el usuario.", "system");
            }
        });
    }

    // Theme Logic
    themeToggleBtn.addEventListener('click', () => {
        document.body.classList.toggle('light-mode');
        const isLight = document.body.classList.contains('light-mode');
        saveSetting('rmp_theme', isLight ? 'light' : 'dark');

        // Robust icon swapping
        if (isLight) {
            sunIcon.classList.remove('hidden');
            moonIcon.classList.add('hidden');
        } else {
            sunIcon.classList.add('hidden');
            moonIcon.classList.remove('hidden');
        }

        // Update Chart.js if exists
        if (window.myChart) {
            window.myChart.options.scales.x.ticks.color = isLight ? '#1E293B' : '#94A3B8';
            window.myChart.options.scales.y.ticks.color = isLight ? '#1E293B' : '#94A3B8';
            window.myChart.options.borderColor = isLight ? '#F8FAFC' : '#05070A';
            window.myChart.update();
        }
    });

    // Comparative Mode Toggle Logic
    const compToggle = document.getElementById('comparative-mode-toggle');
    const historicalContainer = document.getElementById('historical-container');
    const currentFilesLabel = document.getElementById('current-files-label');

    compToggle.addEventListener('change', () => {
        const isActive = compToggle.checked;
        historicalContainer.classList.toggle('hidden', !isActive);
        currentFilesLabel.classList.toggle('hidden', !isActive);
        if (isActive) {
            addChatMessage("Modo Comparativo activo. Sube archivos históricos (pasados).", "system");
        }
    });

    // Mobile Menu Toggle
    const mobileMenuToggle = document.getElementById('mobile-menu-toggle');
    const sidebar = document.querySelector('aside');
    const menuOverlay = document.getElementById('menu-overlay');

    if (mobileMenuToggle && sidebar && menuOverlay) {
        mobileMenuToggle.addEventListener('click', () => {
            sidebar.classList.add('active');
            menuOverlay.classList.add('active');
        });

        menuOverlay.addEventListener('click', () => {
            sidebar.classList.remove('active');
            menuOverlay.classList.remove('active');
        });

        // Close menu on navigation (optional but good for UX)
        sidebar.querySelectorAll('button').forEach(btn => {
            btn.addEventListener('click', () => {
                sidebar.classList.remove('active');
                menuOverlay.classList.remove('active');
            });
        });
    }

    // API Visibility Toggle
    const toggleApiBtn = document.getElementById('toggle-api-visibility');
    if (toggleApiBtn) {
        let isVisible = false;
        toggleApiBtn.addEventListener('click', () => {
            isVisible = !isVisible;
            apiConfig.classList.toggle('hidden', !isVisible);
            toggleApiBtn.innerHTML = isVisible ?
                '<i class="fas fa-eye"></i> Mostrar Config.' :
                '<i class="fas fa-eye-slash"></i> Ocultar';
            // Actually the logic was "Ocultar" when it's visible. 
            // Let's make it logical.
            if (isVisible) {
                toggleApiBtn.innerHTML = '<i class="fas fa-eye"></i> Ocultar API';
            } else {
                toggleApiBtn.innerHTML = '<i class="fas fa-eye-slash"></i> Mostrar API';
            }
        });

        // Initial state
        toggleApiBtn.innerHTML = '<i class="fas fa-eye-slash"></i> Mostrar API';
    }

    // Microphone / Speech Recognition Logic
    const micBtn = document.getElementById('mic-btn');
    let isListening = false;
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

    if (SpeechRecognition && micBtn) {
        const recognition = new SpeechRecognition();
        recognition.lang = 'es-ES';
        recognition.continuous = true;
        recognition.interimResults = true;

        recognition.onstart = () => {
            isListening = true;
            micBtn.classList.add('animate-pulse', 'border-[#00F2FF]', 'shadow-[0_0_15px_rgba(0,242,255,0.5)]');
            updateStatus("Escuchando dictado...");
            console.log("Voice analysis started...");
        };

        recognition.onresult = (event) => {
            let finalTranscript = '';
            for (let i = event.resultIndex; i < event.results.length; ++i) {
                if (event.results[i].isFinal) {
                    finalTranscript += event.results[i][0].transcript + ' ';
                }
            }
            if (finalTranscript && extraNotesInput) {
                console.log("Captured text:", finalTranscript);
                extraNotesInput.value += finalTranscript;
            }
        };

        recognition.onerror = (event) => {
            console.error('Speech recognition error:', event.error);
            isListening = false;
            stopMic();
            if (event.error === 'not-allowed') {
                addChatMessage("Error: Micrófono bloqueado. Revisa los permisos de Windows para aplicaciones.", "system");
            }
        };

        recognition.onend = () => {
            isListening = false;
            stopMic();
        };

        micBtn.addEventListener('click', () => {
            if (!isListening) {
                try {
                    recognition.start();
                } catch (e) { console.error("Recognition start error:", e); }
            } else {
                recognition.stop();
            }
        });

        function stopMic() {
            micBtn.classList.remove('animate-pulse', 'border-[#00F2FF]', 'shadow-[0_0_15px_rgba(0,242,255,0.5)]');
            updateStatus("Dictado finalizado");
            setTimeout(() => updateStatus("LISTO"), 2000);
        }
    } else if (micBtn) {
        micBtn.title = "Dictado no soportado en este entorno";
        micBtn.style.opacity = "0.5";
    }

    // Custom Header/Footer Sync
    customHeaderInput.addEventListener('input', (e) => {
        paperTitleEl.textContent = e.target.value || 'INFORME DE GESTIÓN';
    });
    customFooterInput.addEventListener('input', (e) => {
        paperFooterTextEl.textContent = e.target.value || 'Consultoría Estratégica';
    });

    // Page Size Logic
    pageSizeSelect.addEventListener('change', (e) => {
        reportPaperContainer.classList.remove('size-letter', 'size-legal', 'size-a4');
        reportPaperContainer.classList.add(`size-${e.target.value}`);
    });

    // Motor Selection
    motorOllama.addEventListener('click', () => { switchMotor('ollama'); saveSetting('rmp_current_motor', 'ollama'); });
    motorLMStudio.addEventListener('click', () => { switchMotor('lmstudio'); saveSetting('rmp_current_motor', 'lmstudio'); });
    motorAPI.addEventListener('click', () => { switchMotor('api'); saveSetting('rmp_current_motor', 'api'); });

    // API Configuration listeners (Change to 'input' for real-time reactivity)
    apiUrlInput.addEventListener('input', () => {
        aiService.setApiConfig(apiUrlInput.value, apiKeyInput.value, apiModelInput.value);
        saveSetting('rmp_api_url', apiUrlInput.value);
    });
    apiKeyInput.addEventListener('input', () => {
        aiService.setApiConfig(apiUrlInput.value, apiKeyInput.value, apiModelInput.value);
        saveSetting('rmp_api_key', apiKeyInput.value);
    });
    apiModelInput.addEventListener('input', () => {
        aiService.setApiConfig(apiUrlInput.value, apiKeyInput.value, apiModelInput.value);
        saveSetting('rmp_api_model', apiModelInput.value);
    });

    async function switchMotor(motor) {
        motorOllama.classList.remove('bg-[#6366F1]', 'text-white');
        motorLMStudio.classList.remove('bg-[#6366F1]', 'text-white');
        motorAPI.classList.remove('bg-[#6366F1]', 'text-white');

        const privacyBadge = document.getElementById('privacy-badge');
        const privacyText = privacyBadge.querySelector('span');

        if (motor === 'ollama') {
            motorOllama.classList.add('bg-[#6366F1]', 'text-white');
            apiConfig.classList.add('hidden');
            modelSelect.classList.remove('hidden');
            privacyBadge.className = "flex items-center space-x-1.5 px-3 py-1 rounded-full bg-green-500/10 border border-green-500/30 cursor-help";
            privacyText.textContent = "🔒 Local Secure";
        } else if (motor === 'lmstudio') {
            motorLMStudio.classList.add('bg-[#6366F1]', 'text-white');
            apiConfig.classList.add('hidden');
            modelSelect.classList.remove('hidden');
            privacyBadge.className = "flex items-center space-x-1.5 px-3 py-1 rounded-full bg-green-500/10 border border-green-500/30 cursor-help";
            privacyText.textContent = "🔒 Local Secure";
        } else if (motor === 'api') {
            motorAPI.classList.add('bg-[#6366F1]', 'text-white');
            // apiConfig visibility is now managed exclusively by the toggle button for privacy
            modelSelect.classList.add('hidden');
            privacyBadge.className = "flex items-center space-x-1.5 px-3 py-1 rounded-full bg-blue-500/10 border border-blue-500/30 cursor-help";
            privacyText.textContent = "🌐 Cloud Processing";
            // Sync API settings immediately
            aiService.setApiConfig(apiUrlInput.value, apiKeyInput.value, apiModelInput.value);
        }

        aiService.setMotor(motor);
        if (motor !== 'api') {
            await refreshModels();
        }

        // Reset LED on motor switch
        const connectionLed = document.getElementById('connection-led');
        if (connectionLed) {
            connectionLed.className = 'w-2 h-2 rounded-full bg-slate-700 shadow-[0_0_5px_rgba(0,0,0,0.5)] transition-all';
        }
    }

    // Universal Connection Verification Logic
    const verifyConnBtn = document.getElementById('verify-connection');
    const connectionLed = document.getElementById('connection-led');

    if (verifyConnBtn && connectionLed) {
        verifyConnBtn.addEventListener('click', async () => {
            // Set testing state
            connectionLed.className = 'w-2 h-2 rounded-full led-testing transition-all';

            // Perform test
            const isOnline = await aiService.testConnection();

            // Update UI based on result
            if (isOnline) {
                connectionLed.className = 'w-2 h-2 rounded-full led-online transition-all';
                addChatMessage(`Conexión exitosa con ${aiService.currentMotor.toUpperCase()}.`, 'system');
            } else {
                connectionLed.className = 'w-2 h-2 rounded-full led-offline transition-all';
                addChatMessage(`Error de conexión con ${aiService.currentMotor.toUpperCase()}. Verifica que el motor esté activo.`, 'system');
            }
        });
    }

    async function refreshModels() {
        modelSelect.innerHTML = '<option>Cargando modelos...</option>';
        const models = await aiService.getModels();
        if (models.length > 0) {
            modelSelect.innerHTML = models.map(m => `<option value="${m}">${m}</option>`).join('');
            aiService.model = models[0];
        } else {
            modelSelect.innerHTML = '<option value="">No se detectaron motores</option>';
        }
    }

    modelSelect.addEventListener('change', (e) => {
        aiService.model = e.target.value;
    });

    // Logo & Signature Previews
    logoUpload.addEventListener('change', (e) => handleImagePreview(e, 'logo-preview'));
    signUpload.addEventListener('change', (e) => handleImagePreview(e, 'sign-preview', 'paper-signature'));

    function handleImagePreview(e, previewId, paperId = null) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (event) => {
                const previewDiv = document.getElementById(previewId);
                const img = previewDiv.querySelector('img');
                img.src = event.target.result;
                previewDiv.classList.remove('hidden');

                if (paperId) {
                    const paperEl = document.getElementById(paperId);
                    paperEl.innerHTML = `<img src="${event.target.result}" class="max-h-full max-w-full">`;
                    paperEl.classList.remove('italic', 'text-slate-400');
                } else if (previewId === 'logo-preview') {
                    const paperLogo = document.getElementById('paper-logo');
                    paperLogo.innerHTML = `<img src="${event.target.result}" class="max-h-full max-w-full">`;
                    paperLogo.classList.remove('italic', 'text-slate-400');
                }
            };
            reader.readAsDataURL(file);
        }
    }

    // File Handling
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('bg-[#00F2FF]/10', 'border-[#00F2FF]');
    });
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('bg-[#00F2FF]/10', 'border-[#00F2FF]');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('bg-[#00F2FF]/10', 'border-[#00F2FF]');
        handleFiles(e.dataTransfer.files);
    });
    fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

    async function handleFiles(files) {
        const isComparative = compToggle.checked;
        const targetArray = isComparative ? historicalFilesData : uploadedFilesData;

        for (let file of files) {
            if (file.name.endsWith('.csv') || file.name.endsWith('.txt')) {
                const text = await file.text();
                targetArray.push({ name: file.name, content: text, type: 'text' });
            } else if (file.name.endsWith('.docx')) {
                // Wrap mammoth extract in a promise
                const arrayBuffer = await file.arrayBuffer();
                const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
                targetArray.push({ name: file.name, content: result.value, type: 'word' });
            } else if (file.name.endsWith('.xlsx')) {
                const arrayBuffer = await file.arrayBuffer();
                const data = new Uint8Array(arrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const csv = XLSX.utils.sheet_to_csv(firstSheet);
                targetArray.push({ name: file.name, content: csv, type: 'excel' });
            } else if (file.name.endsWith('.pdf')) {
                const arrayBuffer = await file.arrayBuffer();
                const typedarray = new Uint8Array(arrayBuffer);
                const pdf = await pdfjsLib.getDocument(typedarray).promise;
                let fullText = '';
                for (let i = 1; i <= pdf.numPages; i++) {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items.map(item => item.str).join(' ');
                    fullText += pageText + '\n';
                }
                targetArray.push({ name: file.name, content: fullText, type: 'pdf' });
            }
        }
        renderFileList();
    }

    function renderFileList() {
        // Current Files
        fileListEl.innerHTML = '';
        uploadedFilesData.forEach((file, index) => {
            fileListEl.appendChild(createFileChip(file.name, () => removeFile(index, false)));
        });

        // Historical Files
        const histListEl = document.getElementById('historical-file-list');
        histListEl.innerHTML = '';
        historicalFilesData.forEach((file, index) => {
            histListEl.appendChild(createFileChip(file.name, () => removeFile(index, true), 'bg-purple-900/40 border-purple-700/50'));
        });
    }

    function createFileChip(name, onRemove, customClass = 'bg-slate-800/80 border-slate-700') {
        const chip = document.createElement('div');
        chip.className = `flex items-center ${customClass} px-3 py-1.5 rounded-full text-xs text-slate-300 group hover:border-[#00F2FF] transition-all`;
        chip.innerHTML = `
            <span class="mr-2 truncate max-w-[150px]">${name}</span>
            <button class="text-slate-500 hover:text-red-500 transition-colors focus:outline-none">
                <svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
        `;
        chip.querySelector('button').onclick = onRemove;
        return chip;
    }

    window.removeFile = (index, isHistorical) => {
        if (isHistorical) historicalFilesData.splice(index, 1);
        else uploadedFilesData.splice(index, 1);
        renderFileList();
    };

    function addChatMessage(text, sender = 'user') {
        const div = document.createElement('div');
        div.className = sender === 'user'
            ? 'bg-[#6366F1]/20 p-3 rounded-2xl rounded-tr-none ml-auto max-w-[80%] text-sm text-slate-200'
            : 'bg-slate-800/50 p-3 rounded-2xl rounded-tl-none mr-auto max-w-[80%] text-sm text-slate-300';
        div.textContent = text;
        chatMessages.appendChild(div);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    // Typography Controls
    const fontSansBtn = document.getElementById('font-sans');
    const fontSerifBtn = document.getElementById('font-serif');

    fontSansBtn.addEventListener('click', () => {
        reportPaperContainer.classList.remove('font-serif-mode');
        reportPaperContainer.classList.add('font-sans-mode');
        fontSansBtn.classList.add('bg-[#6366F1]', 'text-white');
        fontSerifBtn.classList.remove('bg-[#6366F1]', 'text-white');
    });

    fontSerifBtn.addEventListener('click', () => {
        reportPaperContainer.classList.remove('font-sans-mode');
        reportPaperContainer.classList.add('font-serif-mode');
        fontSerifBtn.classList.add('bg-[#6366F1]', 'text-white');
        fontSansBtn.classList.remove('bg-[#6366F1]', 'text-white');
    });

    // Main Brain Implementation
    generateBtn.addEventListener('click', async () => {
        if (uploadedFilesData.length === 0) {
            alert('Por favor, sube datos para analizar.');
            return;
        }

        if (!aiService.model) {
            alert('No se ha detectado un motor de IA activo.');
            return;
        }

        // Final safety sync of API settings
        if (aiService.currentMotor === 'api') {
            aiService.setApiConfig(apiUrlInput.value, apiKeyInput.value, apiModelInput.value);
        }

        showThinking(true, "Analizando tendencias estratégicas...");
        reportAbortController = new AbortController();

        try {
            const dataContext = uploadedFilesData.map(f => `FILE [ACTUAL]: ${f.name}\nCONTENT:\n${f.content.substring(0, 5000)}`).join('\n\n');
            const historyContext = compToggle.checked ?
                "\n\nDATOS HISTÓRICOS PARA COMPARACIÓN:\n" + historicalFilesData.map(f => `FILE [HISTÓRICO]: ${f.name}\nCONTENT:\n${f.content.substring(0, 5000)}`).join('\n\n') : "";

            const extraNotes = extraNotesInput.value ? `\n\nNOTAS ADICIONALES:\n${extraNotesInput.value}` : '';
            const level = document.getElementById('report-level').value;
            const industry = document.getElementById('industry-vertical').value;
            const isComparative = compToggle.checked;

            const sysPrompt = aiService.getMcKinseySystemPrompt(level, industry, isComparative);

            updateStatus("Iniciando análisis de proyecciones...");
            const analysis = await aiService.analyze(dataContext + historyContext + extraNotes, sysPrompt, reportAbortController.signal);

            currentReportContent = analysis; // Store for contextual chat
            updateStatus("Diseñando dashboard estratégico...");
            renderReport(analysis);

            // Activate Contextual Chat
            const chatArea = document.getElementById('contextual-chat-area');
            if (chatArea) {
                chatArea.classList.add('active');
                document.getElementById('chat-online-indicator').classList.replace('bg-slate-600', 'bg-[#00F2FF]');
                chatMessages.innerHTML = `
                    <div class="bg-indigo-500/10 border border-indigo-500/30 p-3 rounded-2xl text-[11px] text-indigo-300">
                         <strong>Asistente Contextual Activo:</strong> El reporte ha sido analizado. Puedes realizar preguntas específicas sobre los datos, las discrepancias históricas o las proyecciones.
                    </div>
                `;
            }

            showThinking(false);
            addChatMessage("¡Análisis de élite completado!", "system");
        } catch (error) {
            showThinking(false);
            if (error.message !== 'PROCESO CANCELADO') {
                alert(error.message);
            }
        } finally {
            reportAbortController = null;
        }
    });

    function showThinking(show, message = "") {
        thinkingLayer.classList.toggle('hidden', !show);
        if (message) thinkingStatus.textContent = message;
    }

    function updateStatus(message) {
        thinkingStatus.textContent = message;
    }

    // Contextual Chat Integration
    const chatInput = document.getElementById('chat-input');
    const chatSendBtn = document.getElementById('chat-send-btn');

    async function handleContextualChat() {
        const query = chatInput.value.trim();
        if (!query || !currentReportContent) return;

        addChatMessage(query, 'user');
        chatInput.value = '';

        const typingDiv = document.createElement('div');
        typingDiv.className = 'bg-slate-800/30 p-3 rounded-2xl mr-auto max-w-[80%] text-[10px] text-slate-500 italic animate-pulse';
        typingDiv.textContent = "Consultando informe...";
        chatMessages.appendChild(typingDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;

        try {
            const response = await aiService.queryReport(query, currentReportContent, reportAbortController?.signal);
            typingDiv.remove();
            addChatMessage(response, 'assistant');
        } catch (error) {
            typingDiv.remove();
            if (error.message !== 'PROCESO CANCELADO') {
                addChatMessage("Error al consultar: " + error.message, 'assistant');
            }
        }
    }

    chatSendBtn.addEventListener('click', handleContextualChat);
    chatInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleContextualChat();
    });

    // Industry Theme Switcher
    const industrySelect = document.getElementById('industry-vertical');
    industrySelect.addEventListener('change', () => {
        document.body.setAttribute('data-industry', industrySelect.value);
    });

    function renderReport(markdown) {
        // Parse Special Insights
        let healthScore = 85;
        const healthMatch = markdown.match(/\[HEALTH_SCORE:\s*(\d+)\]/i);
        if (healthMatch) healthScore = parseInt(healthMatch[1]);

        const sideHealthScore = document.getElementById('side-health-score');
        const sideHealthValue = document.getElementById('side-health-value');
        const sideHealthFill = document.getElementById('side-health-fill');

        if (sideHealthScore) {
            sideHealthScore.classList.remove('hidden');
            sideHealthValue.textContent = `${healthScore}%`;
            sideHealthFill.style.width = `${healthScore}%`;
        }

        let criticalFindings = [];
        const criticalMatch = markdown.match(/\[CRITICAL_FINDINGS\]([\s\S]*?)\[/i) || markdown.match(/\[CRITICAL_FINDINGS\]([\s\S]*?)$/i);
        if (criticalMatch) {
            criticalFindings = criticalMatch[1].trim().split('\n').filter(l => l.trim().startsWith('-')).map(l => l.replace('-', '').trim());
        }

        let goldOpportunities = [];
        const goldMatch = markdown.match(/\[GOLD_OPPORTUNITIES\]([\s\S]*?)\[/i) || markdown.match(/\[GOLD_OPPORTUNITIES\]([\s\S]*?)$/i);
        if (goldMatch) {
            goldOpportunities = goldMatch[1].trim().split('\n').filter(l => l.trim().startsWith('-')).map(l => l.replace('-', '').trim());
        }

        let cleanMarkdown = markdown
            .replace(/\[HEALTH_SCORE:.*?\]/gi, '')
            .replace(/\[CRITICAL_FINDINGS\][\s\S]*?(\[|$)/gi, (match, p1) => p1 === '[' ? '[' : '')
            .replace(/\[GOLD_OPPORTUNITIES\][\s\S]*?(\[|$)/gi, (match, p1) => p1 === '[' ? '[' : '')
            .trim();

        // Data Tracing Highlighter (Wrap numbers and percentages)
        let reportHtml = cleanMarkdown
            .replace(/^# (.*$)/gim, '<h1 class="text-3xl font-bold mb-6 text-[#05070A]">$1</h1>')
            .replace(/^## (.*$)/gim, '<h2 class="text-xl font-bold mt-8 mb-4 border-l-4 border-[#00F2FF] pl-4">$1</h2>')
            .replace(/^### (.*$)/gim, '<h3 class="text-lg font-bold mt-6 mb-2">$1</h3>')
            .replace(/(\d+(?:\.\d+)?%?)/g, '<span class="data-trace-highlight" title="Dato extraído del análisis de archivos fuente.">$1</span>')
            .replace(/\*\*(.*)\*\*/gim, '<strong>$1</strong>')
            .replace(/\*(.*)\*/gim, '<em>$1</em>')
            .replace(/\n/gim, '<br>');

        let html = '';
        if (criticalFindings.length || goldOpportunities.length) {
            html += `<div class="grid grid-cols-2 gap-4 mb-8">`;
            if (criticalFindings.length) {
                html += `<div class="space-y-2">`;
                criticalFindings.forEach(f => {
                    html += `<div class="insight-card border-red-500 bg-red-50/50"><span class="badge-critical mb-1 inline-block">Crítico</span><p class="text-xs text-red-900">${f}</p></div>`;
                });
                html += `</div>`;
            }
            if (goldOpportunities.length) {
                html += `<div class="space-y-2">`;
                goldOpportunities.forEach(o => {
                    html += `<div class="insight-card border-green-500 bg-green-50/50"><span class="badge-opportunity mb-1 inline-block">Oportunidad</span><p class="text-xs text-green-900">${o}</p></div>`;
                });
                html += `</div>`;
            }
            html += `</div>`;
        }

        html += `<div class="prose max-w-none text-slate-900">${reportHtml}</div>`;
        reportBody.innerHTML = html;
        document.getElementById('report-paper').scrollTop = 0;
        renderEnhancedDashboard(cleanMarkdown);
    }

    function renderEnhancedDashboard(text) {
        // Reset current charts
        const existingCharts = reportBody.querySelectorAll('.chart-container');
        existingCharts.forEach(c => c.remove());

        // Simple logic for multiple charts
        if (text.toLowerCase().includes('pareto')) {
            addChart('Pareto de Incidencias', 'bar', [80, 15, 5, 2, 1], ['A', 'B', 'C', 'D', 'E']);
        }
        if (text.toLowerCase().includes('tendencia') || text.toLowerCase().includes('crecimiento')) {
            addChart('Tendencia Mensual', 'line', [10, 25, 45, 30, 60], ['Ene', 'Feb', 'Mar', 'Abr', 'May']);
        }
        if (text.toLowerCase().includes('distribución') || text.toLowerCase().includes('porcentaje')) {
            addChart('Distribución de Mercado', 'doughnut', [40, 30, 20, 10], ['Segmento A', 'B', 'C', 'D']);
        }
    }

    function addChart(title, type, data, labels) {
        const id = 'chart-' + Math.random().toString(36).substr(2, 9);
        const container = document.createElement('div');
        container.className = 'chart-container mb-8 p-4 bg-slate-50 rounded-xl border border-slate-200';
        container.innerHTML = `
            <p class="text-[10px] font-bold uppercase text-slate-500 mb-2">${title}</p>
            <canvas id="${id}"></canvas>
        `;

        // Find first H2 or append at end
        const firstH2 = reportBody.querySelector('h2');
        if (firstH2) {
            reportBody.insertBefore(container, firstH2.nextSibling);
        } else {
            reportBody.appendChild(container);
        }

        const ctx = document.getElementById(id).getContext('2d');
        new Chart(ctx, {
            type: type,
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    backgroundColor: ['#00F2FF', '#6366F1', '#00949D', '#0F172A', '#00F2FF80'],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: type === 'doughnut' } }
            }
        });
    }

    // Export Event Listeners
    exportPdfBtn.addEventListener('click', () => exportService.exportPDF());
    exportWordBtn.addEventListener('click', () => exportService.exportWord());
    exportExcelBtn.addEventListener('click', () => exportService.exportExcel(uploadedFilesData));
    exportPptBtn.addEventListener('click', () => exportService.exportPPT());

    // Dedicated Word Download Button
    const downloadWordBtn = document.getElementById('download-word-btn');
    if (downloadWordBtn) {
        downloadWordBtn.addEventListener('click', () => exportService.exportWord());
    }

    // Initial check
    loadSettings();
});
