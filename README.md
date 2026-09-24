
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Institución Educativa La Esperanza — Resultados Saber 11°</title>
    
    <!-- Google Fonts & Font Awesome Icons -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <!-- Chart.js & SheetJS -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

    <style>
        :root {
            --primary: #1e3a8a;
            --primary-light: #3b82f6;
            --bg-body: #f8fafc;
            --card-bg: #ffffff;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --border-color: #e2e8f0;
            
            /* Dynamic Level Colors */
            --level-critico: #ef4444;      /* Rojo */
            --level-bajo: #f59e0b;         /* Amarillo */
            --level-medio: #f97316;        /* Naranja */
            --level-alto: #2563eb;         /* Azul */
            --level-sobresaliente: #10b981;/* Verde */

            /* Subject Colors */
            --subject-lc: #3b82f6;  /* Lectura Crítica - Azul */
            --subject-mat: #8b5cf6; /* Matemáticas - Morado */
            --subject-soc: #f59e0b; /* Sociales - Ámbar */
            --subject-cn: #10b981;  /* Ciencias Naturales - Esmeralda */
            --subject-ing: #ec4899; /* Inglés - Rosa */
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }

        body {
            background-color: var(--bg-body);
            color: var(--text-main);
            padding-bottom: 3rem;
            line-height: 1.5;
        }

        /* HEADER INSTITUCIONAL */
        header {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #2563eb 100%);
            color: #ffffff;
            padding: 2.5rem 2rem 3rem 2rem;
            border-bottom-left-radius: 1.5rem;
            border-bottom-right-radius: 1.5rem;
            box-shadow: 0 10px 25px -5px rgba(30, 58, 138, 0.3);
            margin-bottom: 2rem;
        }

        .header-container {
            max-width: 1280px;
            margin: 0 auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1.5rem;
        }

        .brand-info h1 {
            font-size: 2.2rem;
            font-weight: 800;
            letter-spacing: -0.025em;
            margin-bottom: 0.25rem;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        .brand-info p {
            font-size: 1.1rem;
            color: #93c5fd;
            font-weight: 500;
        }

        .header-actions {
            display: flex;
            align-items: center;
            gap: 1rem;
        }

        .btn-export {
            background: #10b981;
            color: #ffffff;
            border: none;
            padding: 0.75rem 1.5rem;
            border-radius: 0.75rem;
            font-weight: 700;
            font-size: 0.95rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            transition: all 0.2s ease;
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        }

        .btn-export:hover {
            background: #059669;
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(16, 185, 129, 0.4);
        }

        /* MAIN CONTAINER */
        .container {
            max-width: 1280px;
            margin: 0 auto;
            padding: 0 1.5rem;
        }

        /* FILTERS & BAR */
        .controls-card {
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            padding: 1.25rem 1.75rem;
            border-radius: 1rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            margin-bottom: 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1rem;
        }

        .filter-group {
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        .filter-group label {
            font-weight: 700;
            color: var(--text-main);
            font-size: 0.95rem;
        }

        .filter-select {
            padding: 0.6rem 1.2rem;
            border: 2px solid var(--border-color);
            border-radius: 0.6rem;
            font-weight: 600;
            color: var(--text-main);
            background-color: #f8fafc;
            outline: none;
            cursor: pointer;
            transition: border-color 0.2s;
        }

        .filter-select:focus {
            border-color: var(--primary-light);
        }

        /* KPI CARDS */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 1.25rem;
            margin-bottom: 2.5rem;
        }

        .kpi-card {
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 1.25rem;
            padding: 1.5rem;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.03);
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            position: relative;
            overflow: hidden;
            transition: transform 0.2s ease;
        }

        .kpi-card:hover {
            transform: translateY(-4px);
        }

        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 5px;
        }

        .kpi-card.blue::before { background: #3b82f6; }
        .kpi-card.purple::before { background: #8b5cf6; }
        .kpi-card.emerald::before { background: #10b981; }
        .kpi-card.pink::before { background: #ec4899; }
        .kpi-card.amber::before { background: #f59e0b; }

        .kpi-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            color: var(--text-muted);
            font-size: 0.875rem;
            font-weight: 600;
            margin-bottom: 0.75rem;
        }

        .kpi-icon {
            width: 42px;
            height: 42px;
            border-radius: 0.75rem;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.2rem;
        }

        .kpi-card.blue .kpi-icon { background: #eff6ff; color: #3b82f6; }
        .kpi-card.purple .kpi-icon { background: #f3e8ff; color: #8b5cf6; }
        .kpi-card.emerald .kpi-icon { background: #ecfdf5; color: #10b981; }
        .kpi-card.pink .kpi-icon { background: #fdf2f8; color: #ec4899; }
        .kpi-card.amber .kpi-icon { background: #fffbeb; color: #f59e0b; }

        .kpi-value {
            font-size: 2.25rem;
            font-weight: 800;
            color: var(--text-main);
            letter-spacing: -0.03em;
        }

        .kpi-subtext {
            font-size: 0.8rem;
            color: var(--text-muted);
            margin-top: 0.25rem;
            font-weight: 500;
        }

        /* SECTION TITLES */
        .section-title {
            font-size: 1.35rem;
            font-weight: 800;
            color: var(--text-main);
            margin-bottom: 1.25rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }

        /* CHARTS SECTION */
        .charts-grid {
            display: grid;
            grid-template-columns: 3fr 2fr;
            gap: 1.5rem;
            margin-bottom: 2.5rem;
        }

        @media (max-width: 992px) {
            .charts-grid {
                grid-template-columns: 1fr;
            }
        }

        .chart-card {
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 1.25rem;
            padding: 1.75rem;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.03);
            display: flex;
            flex-direction: column;
        }

        .chart-header {
            margin-bottom: 1.25rem;
        }

        .chart-header h3 {
            font-size: 1.1rem;
            font-weight: 700;
            color: var(--text-main);
        }

        .chart-container {
            position: relative;
            flex-grow: 1;
            min-height: 280px;
        }

        /* DAFO MATRIX */
        .dafo-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 1.25rem;
            margin-bottom: 2.5rem;
        }

        @media (max-width: 768px) {
            .dafo-grid {
                grid-template-columns: 1fr;
            }
        }

        .dafo-card {
            border-radius: 1.25rem;
            padding: 1.5rem;
            border: 1px solid transparent;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02);
            transition: transform 0.2s;
        }

        .dafo-card:hover {
            transform: translateY(-2px);
        }

        .dafo-card h4 {
            font-size: 1.1rem;
            font-weight: 800;
            margin-bottom: 0.75rem;
            display: flex;
            align-items: center;
            gap: 0.6rem;
        }

        .dafo-card ul {
            list-style: none;
            padding-left: 0;
        }

        .dafo-card li {
            font-size: 0.9rem;
            margin-bottom: 0.5rem;
            position: relative;
            padding-left: 1.25rem;
        }

        .dafo-card li::before {
            content: '•';
            position: absolute;
            left: 0;
            font-weight: bold;
            font-size: 1.2rem;
            line-height: 1;
        }

        /* DAFO Colors */
        .dafo-fortalezas { background-color: #f0fdf4; border-color: #bbf7d0; color: #166534; }
        .dafo-fortalezas h4 { color: #15803d; }
        .dafo-fortalezas li::before { color: #15803d; }

        .dafo-debilidades { background-color: #fef2f2; border-color: #fecaca; color: #991b1b; }
        .dafo-debilidades h4 { color: #dc2626; }
        .dafo-debilidades li::before { color: #dc2626; }

        .dafo-oportunidades { background-color: #eff6ff; border-color: #bfdbfe; color: #1e40af; }
        .dafo-oportunidades h4 { color: #2563eb; }
        .dafo-oportunidades li::before { color: #2563eb; }

        .dafo-amenazas { background-color: #fffbeb; border-color: #fde68a; color: #92400e; }
        .dafo-amenazas h4 { color: #d97706; }
        .dafo-amenazas li::before { color: #d97706; }

        /* RANKING TABLE */
        .table-card {
            background: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 1.25rem;
            padding: 1.5rem;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.03);
            overflow: hidden;
        }

        .table-responsive {
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 0.9rem;
        }

        th {
            background-color: #f8fafc;
            color: var(--text-muted);
            font-weight: 700;
            padding: 1rem 0.75rem;
            border-bottom: 2px solid var(--border-color);
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.05em;
        }

        td {
            padding: 1rem 0.75rem;
            border-bottom: 1px solid var(--border-color);
            color: var(--text-main);
            font-weight: 500;
        }

        tr:hover {
            background-color: #f8fafc;
        }

        .rank-badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            background-color: #e2e8f0;
            color: var(--text-main);
            font-weight: 800;
            font-size: 0.85rem;
        }

        .rank-top {
            background: linear-gradient(135deg, #f59e0b, #d97706);
            color: white;
        }

        .student-name {
            font-weight: 700;
            color: var(--text-main);
        }

        /* BADGES FOR PERFORMANCE LEVELS */
        .level-badge {
            display: inline-block;
            padding: 0.35rem 0.75rem;
            border-radius: 2rem;
            font-size: 0.75rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.025em;
        }

        .badge-critico { background-color: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }
        .badge-bajo { background-color: #fffbeb; color: #d97706; border: 1px solid #fde68a; }
        .badge-medio { background-color: #fff7ed; color: #ea580c; border: 1px solid #ffedd5; }
        .badge-alto { background-color: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; }
        .badge-sobresaliente { background-color: #ecfdf5; color: #059669; border: 1px solid #a7f3d0; }

        .score-highlight {
            font-weight: 800;
            font-size: 1rem;
            color: var(--primary);
        }

        footer {
            text-align: center;
            margin-top: 3rem;
            color: var(--text-muted);
            font-size: 0.85rem;
        }
    </style>
</head>
<body>

    <!-- ENCABEZADO Y MARCA -->
    <header>
        <div class="header-container">
            <div class="brand-info">
                <h1><i class="fa-solid => fa-school"></i> Institución Educativa La Esperanza</h1>
                <p>Resultados Saber 11° — Oficial Institución Educativa La Esperanza</p>
            </div>
            <div class="header-actions">
                <button class="btn-export" onclick="exportToExcel()">
                    <i class="fa-solid fa-file-excel"></i> Exportar Excel
                </button>
            </div>
        </div>
    </header>

    <div class="container">

        <!-- FILTRO INTERACTIVO -->
        <div class="controls-card">
            <div class="filter-group">
                <i class="fa-solid fa-filter" style="color: var(--primary-light);"></i>
                <label for="gradeFilter">Filtrar por Grado:</label>
                <select id="gradeFilter" class="filter-select" onchange="filterData()">
                    <option value="ALL">Todos los Grados</option>
                    <option value="11°">Grado 11°</option>
                </select>
            </div>
            <div style="font-size: 0.85rem; color: var(--text-muted); font-weight: 600;">
                <i class="fa-solid fa-circle-check" style="color: #10b981;"></i> Datos Oficiales Verificados
            </div>
        </div>

        <!-- INDICADORES PRINCIPALES (KPIs) -->
        <div class="kpi-grid">
            <div class="kpi-card blue">
                <div class="kpi-header">
                    <span>PROMEDIO GLOBAL</span>
                    <div class="kpi-icon"><i class="fa-solid fa-chart-line"></i></div>
                </div>
                <div class="kpi-value" id="kpiAvgGlobal">273</div>
                <div class="kpi-subtext">Puntaje sobre 500 pts</div>
            </div>

            <div class="kpi-card purple">
                <div class="kpi-header">
                    <span>DESVIACIÓN ESTÁNDAR</span>
                    <div class="kpi-icon"><i class="fa-solid fa-arrows-left-right"></i></div>
                </div>
                <div class="kpi-value" id="kpiStdDev">52</div>
                <div class="kpi-subtext">Dispersión de puntajes</div>
            </div>

            <div class="kpi-card emerald">
                <div class="kpi-header">
                    <span>MEJOR ÁREA</span>
                    <div class="kpi-icon"><i class="fa-solid fa-trophy"></i></div>
                </div>
                <div class="kpi-value" style="font-size: 1.5rem;" id="kpiBestArea">C. Naturales</div>
                <div class="kpi-subtext" id="kpiBestAreaScore">Promedio: 60 pts</div>
            </div>

            <div class="kpi-card pink">
                <div class="kpi-header">
                    <span>ÁREA A REFORZAR</span>
                    <div class="kpi-icon"><i class="fa-solid fa-triangle-exclamation"></i></div>
                </div>
                <div class="kpi-value" style="font-size: 1.5rem;" id="kpiWeakArea">Inglés</div>
                <div class="kpi-subtext" id="kpiWeakAreaScore">Promedio: 45 pts</div>
            </div>

            <div class="kpi-card amber">
                <div class="kpi-header">
                    <span>TOTAL EVALUADOS</span>
                    <div class="kpi-icon"><i class="fa-solid fa-users"></i></div>
                </div>
                <div class="kpi-value" id="kpiTotalStudents">11</div>
                <div class="kpi-subtext">Estudiantes evaluados</div>
            </div>
        </div>

        <!-- VISUALIZACIÓN GRÁFICA (Chart.js) -->
        <div class="charts-grid">
            <div class="chart-card">
                <div class="chart-header">
                    <h3><i class="fa-solid fa-chart-bar" style="color: var(--primary-light);"></i> Promedio de Puntaje por Áreas de Evaluación</h3>
                </div>
                <div class="chart-container">
                    <canvas id="barChartAreas"></canvas>
                </div>
            </div>

            <div class="chart-card">
                <div class="chart-header">
                    <h3><i class="fa-solid fa-chart-pie" style="color: var(--primary-light);"></i> Distribución por Niveles de Desempeño</h3>
                </div>
                <div class="chart-container">
                    <canvas id="pieChartLevels"></canvas>
                </div>
            </div>
        </div>

        <!-- MATRIZ DE DIAGNÓSTICO DAFO -->
        <h2 class="section-title"><i class="fa-solid fa-square-poll-vertical" style="color: var(--primary-light);"></i> Matriz de Diagnóstico DAFO</h2>
        <div class="dafo-grid">
            <div class="dafo-card dafo-fortalezas">
                <h4><i class="fa-solid fa-circle-check"></i> Fortalezas (Verde)</h4>
                <ul>
                    <li>Excelente rendimiento sobresaliente en Ciencias Naturales (60 pts) y Matemáticas (59 pts).</li>
                    <li>Presencia de alto desempeño individual, destacando un puntaje superior de 395 pts.</li>
                    <li>27% de los estudiantes alcanzaron el Nivel Alto (300-400 pts).</li>
                </ul>
            </div>

            <div class="dafo-card dafo-debilidades">
                <h4><i class="fa-solid fa-circle-xmark"></i> Debilidades (Rojo)</h4>
                <ul>
                    <li>Bajo promedio general en el área de Inglés (45 pts) y Sociales y Ciudadanas (50 pts).</li>
                    <li>El 36% de los estudiantes se ubican en Nivel Bajo (200-250 pts).</li>
                    <li>Brecha significativa de 183 puntos entre el rendimiento máximo y el mínimo.</li>
                </ul>
            </div>

            <div class="dafo-card dafo-oportunidades">
                <h4><i class="fa-solid fa-lightbulb"></i> Oportunidades (Azul)</h4>
                <ul>
                    <li>36% de la cohorte se halla en Nivel Medio (250-300 pts), con alto potencial para migrar a Nivel Alto.</li>
                    <li>Implementación de talleres focalizados en Lectura Crítica y Competencias Ciudadanas.</li>
                    <li>Aprovechamiento de los estudiantes con altos resultados como tutores pares.</li>
                </ul>
            </div>

            <div class="dafo-card dafo-amenazas">
                <h4><i class="fa-solid fa-triangle-exclamation"></i> Amenazas (Amarillo)</h4>
                <ul>
                    <li>Alta variabilidad (Desviación Estándar = 52) que afecta la consistencia del rendimiento global.</li>
                    <li>Riesgo de estancamiento en la competencia bilingüe (Inglés) si no se ajusta la intensidad horaria.</li>
                    <li>Afectación del índice sintético institucional por concentración en niveles bajos.</li>
                </ul>
            </div>
        </div>

        <!-- TABLA DE RANKING -->
        <h2 class="section-title"><i class="fa-solid fa-list-ol" style="color: var(--primary-light);"></i> Tabla de Ranking de Estudiantes</h2>
        <div class="table-card">
            <div class="table-responsive">
                <table id="rankingTable">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Estudiante</th>
                            <th>Grado</th>
                            <th>Lectura Crítica</th>
                            <th>Matemáticas</th>
                            <th>Sociales</th>
                            <th>C. Naturales</th>
                            <th>Inglés</th>
                            <th>Puntaje Global</th>
                            <th>Nivel de Desempeño</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                        <!-- Generado dinámicamente -->
                    </tbody>
                </table>
            </div>
        </div>

    </div>

    <footer>
        <p>© 2026 Institución Educativa La Esperanza — Plataforma de Gestión de Resultados Saber 11°</p>
    </footer>

    <script>
        // DATOS OFICIALES DE ESTUDIANTES (EXTRAÍDOS DEL EXCEL OFICIAL)
        // Regla obligatoria de redondeo aplicada:
        // Si la parte decimal es > 0,5 se redondea al entero siguiente; si es <= 0,5 se conserva el entero inferior.
        
        function customRound(val) {
            if (val === null || val === undefined) return 0;
            let floor = Math.floor(val);
            let frac = val - floor;
            return frac > 0.5 ? Math.ceil(val) : floor;
        }

        const rawStudentsData = [
            { id: 11, name: "BERROCAL ARRIETA LUIS MATEO", grade: "11°", lc: 74, mat: 80, soc: 67, cn: 100, ing: 64, globalRaw: 395.000000 },
            { id: 34, name: "APARICIO MELENDREZ ANGELA MARÍA", grade: "11°", lc: 59, mat: 69, soc: 53, cn: 68, ing: 45, globalRaw: 304.615385 },
            { id: 12, name: "RUIZ RODRIGUEZ VALENTINA", grade: "11°", lc: 59, mat: 62, soc: 58, cn: 65, ing: 51, globalRaw: 301.153846 },
            { id: 16, name: "MARIMON TUIRAN SARA ENA", grade: "11°", lc: 51, mat: 66, soc: 55, cn: 64, ing: 46, globalRaw: 290.000000 },
            { id: 14, name: "ARRIETA GONZALEZ LORENA", grade: "11°", lc: 55, mat: 65, soc: 53, cn: 59, ing: 51, globalRaw: 287.307692 },
            { id: 13, name: "MENDEZ SIERRA EIDY LUZ", grade: "11°", lc: 52, mat: 65, soc: 37, cn: 67, ing: 45, globalRaw: 272.307692 },
            { id: 8,  name: "JARAMILLO CAMAÑO JUAN PABLO", grade: "11°", lc: 51, mat: 51, soc: 47, cn: 57, ing: 45, globalRaw: 255.000000 },
            { id: 2,  name: "GASPAR SIERRA EMILY VALENTINA", grade: "11°", lc: 55, mat: 44, soc: 44, cn: 50, ing: 40, globalRaw: 238.076923 },
            { id: 15, name: "RUIZ ESTRADA JUAN DAVID", grade: "11°", lc: 43, mat: 43, soc: 55, cn: 40, ing: 40, globalRaw: 224.230769 },
            { id: 19, name: "ORTIZ MADERA CAMILO ANDRES", grade: "11°", lc: 41, mat: 48, soc: 41, cn: 48, ing: 40, globalRaw: 220.769231 },
            { id: 5,  name: "MORENO SANCHEZ EMANUEL", grade: "11°", lc: 34, mat: 53, soc: 40, cn: 47, ing: 29, globalRaw: 211.923077 }
        ];

        // Procesar puntajes de estudiantes aplicando redondeo
        const students = rawStudentsData.map(st => {
            const globalRounded = customRound(st.globalRaw);
            let level = "";
            let badgeClass = "";

            // NIVELES DE DESEMPEÑO SEGÚN RANGOS ESPECÍFICOS
            if (globalRounded < 200) {
                level = "Crítico";
                badgeClass = "badge-critico";
            } else if (globalRounded >= 200 && globalRounded < 250) {
                level = "Bajo";
                badgeClass = "badge-bajo";
            } else if (globalRounded >= 250 && globalRounded < 300) {
                level = "Medio";
                badgeClass = "badge-medio";
            } else if (globalRounded >= 300 && globalRounded < 400) {
                level = "Alto";
                badgeClass = "badge-alto";
            } else {
                level = "Sobresaliente";
                badgeClass = "badge-sobresaliente";
            }

            return {
                ...st,
                lc: customRound(st.lc),
                mat: customRound(st.mat),
                soc: customRound(st.soc),
                cn: customRound(st.cn),
                ing: customRound(st.ing),
                globalScore: globalRounded,
                level: level,
                badgeClass: badgeClass
            };
        });

        // Variables globales para gráficos
        let barChart = null;
        let pieChart = null;

        document.addEventListener("DOMContentLoaded", () => {
            filterData();
        });

        function filterData() {
            const selectedGrade = document.getElementById("gradeFilter").value;
            let filtered = students;

            if (selectedGrade !== "ALL") {
                filtered = students.filter(s => s.grade === selectedGrade);
            }

            // Ordenar de Mayor a Menor por Puntaje Global
            filtered.sort((a, b) => b.globalScore - a.globalScore);

            updateKPIsAndCharts(filtered);
            renderTable(filtered);
        }

        function updateKPIsAndCharts(dataList) {
            if (dataList.length === 0) return;

            // Calculations
            const count = dataList.length;
            const sumGlobal = dataList.reduce((acc, s) => acc + s.globalScore, 0);
            const avgGlobal = customRound(sumGlobal / count);

            // Standard Deviation
            const mean = sumGlobal / count;
            const variance = dataList.reduce((acc, s) => acc + Math.pow(s.globalScore - mean, 2), 0) / (count > 1 ? count - 1 : 1);
            const stdDev = customRound(Math.sqrt(variance));

            // Area Averages
            const avgLC = customRound(dataList.reduce((acc, s) => acc + s.lc, 0) / count);
            const avgMAT = customRound(dataList.reduce((acc, s) => acc + s.mat, 0) / count);
            const avgSOC = customRound(dataList.reduce((acc, s) => acc + s.soc, 0) / count);
            const avgCN = customRound(dataList.reduce((acc, s) => acc + s.cn, 0) / count);
            const avgING = customRound(dataList.reduce((acc, s) => acc + s.ing, 0) / count);

            const areaScores = [
                { name: "Lectura Crítica", score: avgLC },
                { name: "Matemáticas", score: avgMAT },
                { name: "Sociales", score: avgSOC },
                { name: "C. Naturales", score: avgCN },
                { name: "Inglés", score: avgING }
            ];

            areaScores.sort((a, b) => b.score - a.score);
            const bestArea = areaScores[0];
            const weakArea = areaScores[areaScores.length - 1];

            // Update KPI HTML
            document.getElementById("kpiAvgGlobal").innerText = avgGlobal;
            document.getElementById("kpiStdDev").innerText = stdDev;
            document.getElementById("kpiBestArea").innerText = bestArea.name;
            document.getElementById("kpiBestAreaScore").innerText = `Promedio: ${bestArea.score} pts`;
            document.getElementById("kpiWeakArea").innerText = weakArea.name;
            document.getElementById("kpiWeakAreaScore").innerText = `Promedio: ${weakArea.score} pts`;
            document.getElementById("kpiTotalStudents").innerText = count;

            // Level Distribution
            const levelCounts = { "Crítico": 0, "Bajo": 0, "Medio": 0, "Alto": 0, "Sobresaliente": 0 };
            dataList.forEach(s => {
                if (levelCounts[s.level] !== undefined) {
                    levelCounts[s.level]++;
                }
            });

            // Render Charts
            renderBarChart([avgLC, avgMAT, avgSOC, avgCN, avgING]);
            renderPieChart([
                levelCounts["Crítico"],
                levelCounts["Bajo"],
                levelCounts["Medio"],
                levelCounts["Alto"],
                levelCounts["Sobresaliente"]
            ], count);
        }

        function renderBarChart(areaAverages) {
            const ctx = document.getElementById("barChartAreas").getContext("2d");
            
            if (barChart) barChart.destroy();

            barChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: ['Lectura Crítica', 'Matemáticas', 'Sociales', 'C. Naturales', 'Inglés'],
                    datasets: [{
                        label: 'Promedio por Área',
                        data: areaAverages,
                        backgroundColor: [
                            '#3b82f6', // Lectura Crítica - Azul
                            '#8b5cf6', // Matemáticas - Morado
                            '#f59e0b', // Sociales - Ámbar
                            '#10b981', // C. Naturales - Verde Esmeralda
                            '#ec4899'  // Inglés - Rosa
                        ],
                        borderRadius: 8,
                        borderWidth: 0
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    return ` Promedio: ${context.parsed.y} pts`;
                                }
                            }
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            max: 100,
                            ticks: {
                                precision: 0,
                                font: { family: 'Plus Jakarta Sans', weight: '600' }
                            },
                            grid: { color: '#f1f5f9' }
                        },
                        x: {
                            ticks: { font: { family: 'Plus Jakarta Sans', weight: '600' } },
                            grid: { display: false }
                        }
                    }
                }
            });
        }

        function renderPieChart(levelCounts, total) {
            const ctx = document.getElementById("pieChartLevels").getContext("2d");

            if (pieChart) pieChart.destroy();

            pieChart = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: ['Crítico (<200)', 'Bajo (200-249)', 'Medio (250-299)', 'Alto (300-399)', 'Sobresaliente (≥400)'],
                    datasets: [{
                        data: levelCounts,
                        backgroundColor: [
                            '#ef4444', // Crítico - Rojo
                            '#f59e0b', // Bajo - Amarillo/Ámbar
                            '#f97316', // Medio - Naranja
                            '#2563eb', // Alto - Azul
                            '#10b981'  // Sobresaliente - Verde
                        ],
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                font: { family: 'Plus Jakarta Sans', size: 11, weight: '600' },
                                boxWidth: 12,
                                padding: 12
                            }
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const val = context.parsed;
                                    const pct = total > 0 ? customRound((val / total) * 100) : 0;
                                    return ` ${context.label}: ${val} est. (${pct}%)`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function renderTable(dataList) {
            const tbody = document.getElementById("tableBody");
            tbody.innerHTML = "";

            dataList.forEach((s, index) => {
                const tr = document.createElement("tr");
                
                const rankClass = index < 3 ? 'rank-top' : '';

                tr.innerHTML = `
                    <td><span class="rank-badge ${rankClass}">${index + 1}</span></td>
                    <td class="student-name">${s.name}</td>
                    <td>${s.grade}</td>
                    <td>${s.lc}</td>
                    <td>${s.mat}</td>
                    <td>${s.soc}</td>
                    <td>${s.cn}</td>
                    <td>${s.ing}</td>
                    <td class="score-highlight">${s.globalScore}</td>
                    <td><span class="level-badge ${s.badgeClass}">${s.level}</span></td>
                `;
                tbody.appendChild(tr);
            });
        }

        // FUNCIONALIDAD DE EXPORTACIÓN CON SHEETJS
        function exportToExcel() {
            const selectedGrade = document.getElementById("gradeFilter").value;
            let filtered = students;

            if (selectedGrade !== "ALL") {
                filtered = students.filter(s => s.grade === selectedGrade);
            }

            filtered.sort((a, b) => b.globalScore - a.globalScore);

            const exportData = filtered.map((s, index) => ({
                "Ranking": index + 1,
                "Nombre y Apellidos": s.name,
                "Grado": s.grade,
                "Lectura Crítica": s.lc,
                "Matemáticas": s.mat,
                "Sociales y C. Ciudadanas": s.soc,
                "Ciencias Naturales": s.cn,
                "Inglés": s.ing,
                "Puntaje Global": s.globalScore,
                "Nivel de Desempeño": s.level
            }));

            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Resultados Saber 11");

            // Auto-ajustar ancho de columnas
            const max_width = exportData.reduce((w, r) => Math.max(w, r["Nombre y Apellidos"].length), 10);
            worksheet['!cols'] = [
                { wch: 8 },  // Ranking
                { wch: max_width + 5 }, // Nombre
                { wch: 8 },  // Grado
                { wch: 15 }, // LC
                { wch: 15 }, // MAT
                { wch: 22 }, // SOC
                { wch: 18 }, // CN
                { wch: 10 }, // ING
                { wch: 15 }, // Global
                { wch: 18 }  // Nivel
            ];

            XLSX.writeFile(workbook, `Resultados_Saber11_IE_La_Esperanza_${selectedGrade}.xlsx`);
        }
    </script>
</body>
</html>
