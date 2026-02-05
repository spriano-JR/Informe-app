/**
 * REPORT-MASTER PRO - AI Service
 * Manages connectivity with local AI engines (Ollama / LM Studio)
 */

class AIService {
    constructor() {
        this.currentMotor = 'ollama'; // default
        this.baseUrl = 'http://127.0.0.1:11434/api'; // default ollama
        this.model = '';
        this.apiKey = '';
        this.apiUrl = '';
        this.apiModel = '';
    }

    setMotor(motor) {
        this.currentMotor = motor;
        if (motor === 'ollama') {
            this.baseUrl = 'http://127.0.0.1:11434/api';
        } else if (motor === 'lmstudio') {
            this.baseUrl = 'http://127.0.0.1:1234/v1';
        } else if (motor === 'api') {
            this.baseUrl = this.apiUrl || '';
        }
    }

    setApiConfig(url, key, model) {
        this.apiUrl = url;
        this.apiKey = key;
        this.apiModel = model;
        if (this.currentMotor === 'api') {
            this.baseUrl = url;
            this.model = model;
        }
    }

    async getModels() {
        try {
            if (this.currentMotor === 'ollama') {
                const response = await fetch(`${this.baseUrl}/tags`);
                const data = await response.json();
                return data.models.map(m => m.name);
            } else {
                // LM Studio
                const response = await fetch(`${this.baseUrl}/models`);
                const data = await response.json();
                return data.data.map(m => m.id);
            }
        } catch (error) {
            console.error('Error fetching models:', error);
            return [];
        }
    }

    async analyze(prompt, systemPrompt, signal) {
        try {
            const body = this.currentMotor === 'ollama' ?
                { model: this.model, prompt: prompt, system: systemPrompt, stream: false } :
                { model: this.model, messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: prompt }], temperature: 0.7 };

            const url = this.currentMotor === 'ollama' ? `${this.baseUrl}/generate` : `${this.baseUrl}/chat/completions`;

            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', ...(this.currentMotor === 'api' ? { 'Authorization': `Bearer ${this.apiKey}` } : {}) },
                signal: signal,
                body: JSON.stringify(body)
            });

            if (!response.ok) {
                let errorMsg = `Error ${response.status}`;
                try {
                    const errorData = await response.json();
                    errorMsg = errorData.error?.message || errorData.error || errorMsg;
                } catch (e) { }
                throw new Error(errorMsg);
            }

            const data = await response.json();
            return this.currentMotor === 'ollama' ? data.response : data.choices[0].message.content;
        } catch (error) {
            if (error.name === 'AbortError') throw new Error('PROCESO CANCELADO');
            console.error('AI Analysis error:', error);
            throw new Error(`Error de IA: ${error.message}`);
        }
    }

    async testConnection() {
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 5000); // 5s timeout

            let url = "";
            if (this.currentMotor === 'ollama') {
                url = `${this.baseUrl}/tags`;
            } else if (this.currentMotor === 'lmstudio') {
                url = `${this.baseUrl}/models`;
            } else if (this.currentMotor === 'api') {
                url = `${this.baseUrl}/models`;
            }

            const response = await fetch(url, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    ...(this.currentMotor === 'api' ? { 'Authorization': `Bearer ${this.apiKey}` } : {})
                },
                signal: controller.signal
            });

            clearTimeout(timeoutId);
            return response.ok;
        } catch (error) {
            console.error('Connection test failed:', error);
            return false;
        }
    }

    getMcKinseySystemPrompt(level, industry = 'general', isComparative = false) {
        const base = `Eres un Socio Principal de McKinsey especializado en análisis estratégico.`;

        const industries = {
            'general': 'Enfócate en eficiencia global y salud organizacional.',
            'finance': 'Foco crítico en EBITDA, Flujo de Caja, ROE y solvencia. Usa terminología financiera de élite.',
            'sales': 'Foco en CAC, LTV, Tasa de Conversión y Cuota de Mercado. Analiza el embudo de ventas.',
            'ops': 'Foco en Tiempos de Ciclo, Cuellos de Botella, OEE y Cadena de Suministro.'
        };

        const levels = {
            'basic': "Resumen ejecutivo para alta gerencia (Quick Wins).",
            'medium': "Informe técnico y estadístico. Tendencias clave.",
            'pro': "Estrategia profunda. Pareto (80/20), anomalías y acciones tácticas.",
            'ntu': "INFORME NTU: Precisión operacional extrema, organización estructural de datos y hoja de ruta clara."
        };

        const comparativeNote = isComparative ?
            "\nIMPORTANTE: Estás en MODO COMPARATIVO. Analiza el 'Archivo Histórico' vs el 'Archivo Actual'. Calcula el DELTA de rendimiento, identifica si hay progreso o retrocesos y explica la causa raíz." : "";

        return `${base} ${levels[level]} Industria: ${industries[industry]} ${comparativeNote}
        
        MANDATORIO:
        1. Comienza con [HEALTH_SCORE: X] (1-100).
        2. Sigue con [CRITICAL_FINDINGS] y [GOLD_OPPORTUNITIES].
        3. INCLUYE UNA SECCIÓN DE 'PROYECCIONES PREDICTIVAS' (Forecasting): Estima el rendimiento a 3, 6 y 12 meses basado en las tendencias actuales.
        4. Si detectas una cifra anómala, menciónala explícitamente citando el valor.
        
        Usa un tono empresarial de élite en español. Organiza con títulos (##) y listas.`;
    }

    async queryReport(userQuestion, reportContent, signal) {
        const sysPrompt = `Eres un analista experto. Responde preguntas específicas basándote ÚNICAMENTE en el informe proporcionado. Sé breve y preciso. Si la información no está en el informe, indícalo educadamente.`;
        const prompt = `INFORME:\n${reportContent}\n\nPREGUNTA DEL USUARIO: ${userQuestion}`;
        return this.analyze(prompt, sysPrompt, signal);
    }
}

// Export default instance
const aiService = new AIService();
