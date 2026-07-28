/**
 * ============================================
 * DASHBOARD RH 2026 - MAIN JAVASCRIPT
 * ============================================
 * Chemin données: data/data.json
 * Déploiement: https://portailrh/Agirh/Dashboard
 * ============================================
 */

// Configuration globale
const CONFIG = {
    dataPath: 'data/data.json',
    itemsPerPage: 20,
    chartColors: {
        orange: '#FF6B35',
        orangeLight: 'rgba(255, 107, 53, 0.6)',
        blue: '#1E3A5F',
        blueLight: 'rgba(30, 58, 95, 0.6)',
        teal: '#00B4A0',
        tealLight: 'rgba(0, 180, 160, 0.6)',
        purple: '#6B5BFF',
        purpleLight: 'rgba(107, 91, 255, 0.6)',
        gold: '#F59E0B',
        success: '#10B981',
        danger: '#EF4444',
        gray: '#94A3B8',
        grayLight: 'rgba(148, 163, 184, 0.3)'
    }
};

// État global de l'application
let state = {
    data: null,
    filteredEffectif: [],
    filteredMS: [],
    filteredRecrutements: [],
    filteredDeparts: [],
    currentPage: 1,
    filters: {
        etb: '',
        classification: '',
        contrat: '',
        search: ''
    },
    charts: {}
};

// ============================================
// INITIALISATION
// ============================================

document.addEventListener('DOMContentLoaded', async () => {
    await loadData();
    initializeApp();
});

async function loadData() {
    try {
        const response = await fetch(CONFIG.dataPath);
        if (!response.ok) throw new Error('Erreur de chargement des données');
        state.data = await response.json();
        
        // Initialiser les données filtrées
        state.filteredEffectif = [...state.data.effectif];
        state.filteredMS = [...state.data.masseSalariale];
        state.filteredRecrutements = [...state.data.recrutements];
        state.filteredDeparts = [...state.data.departs];
        
        console.log('✅ Données chargées:', {
            effectif: state.data.effectif.length,
            ms: state.data.masseSalariale.length,
            recrutements: state.data.recrutements.length,
            departs: state.data.departs.length
        });
        
        // Mettre à jour la date de mise à jour
        if (state.data.config && state.data.config.lastUpdate) {
            const updateEl = document.getElementById('lastUpdate');
            if (updateEl) {
                updateEl.textContent = formatDateTime(state.data.config.lastUpdate);
            }
        }
    } catch (error) {
        console.error('Erreur:', error);
        showError('Impossible de charger les données. Vérifiez que le fichier data/data.json existe.');
    }
}

function initializeApp() {
    if (!state.data) return;
    
    // Initialiser les graphiques
    initCharts();
    
    // Rendre les tableaux
    renderEffectifTable();
    renderMSTable();
    renderMouvementsTable();
    
    // Calculer et afficher les KPIs
    updateAllKPIs();
    
    // Initialiser les événements
    initEventListeners();
}

function initEventListeners() {
    // Filtres
    document.getElementById('filterETB')?.addEventListener('change', applyFilters);
    document.getElementById('filterClassification')?.addEventListener('change', applyFilters);
    document.getElementById('filterContrat')?.addEventListener('change', applyFilters);
    document.getElementById('globalSearch')?.addEventListener('input', debounce(applyFilters, 300));
    
    // Navigation
    document.querySelectorAll('.nav-link[data-page]').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            showPage(link.dataset.page);
        });
    });
    
    // Bouton actualiser
    document.getElementById('btnRefresh')?.addEventListener('click', refreshData);
}

// ============================================
// NAVIGATION
// ============================================

function showPage(pageId) {
    // Masquer toutes les pages
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    
    // Désactiver tous les liens
    document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
    
    // Afficher la page sélectionnée
    const page = document.getElementById(`page-${pageId}`);
    if (page) page.classList.add('active');
    
    // Activer le lien correspondant
    const link = document.querySelector(`.nav-link[data-page="${pageId}"]`);
    if (link) link.classList.add('active');
    
    // Mettre à jour le titre
    const titles = {
        'synthese': { title: 'Synthèse RH', subtitle: 'Vue d\'ensemble des indicateurs clés' },
        'effectifs': { title: 'Gestion des Effectifs', subtitle: 'Analyse détaillée du capital humain' },
        'masse-salariale': { title: 'Masse Salariale', subtitle: 'Analyse financière et budgétaire' },
        'mouvements': { title: 'Mouvements RH', subtitle: 'Recrutements, départs et indicateurs de rotation' }
    };
    
    if (titles[pageId]) {
        document.getElementById('pageTitle').textContent = titles[pageId].title;
        document.getElementById('pageSubtitle').textContent = titles[pageId].subtitle + ' - ' + (state.data?.config?.moisCourant || '') + ' ' + (state.data?.config?.annee || '');
    }
    
    // Redimensionner les graphiques
    setTimeout(() => {
        Object.values(state.charts).forEach(chart => {
            if (chart && typeof chart.resize === 'function') {
                chart.resize();
            }
        });
    }, 100);
}

// ============================================
// FILTRES
// ============================================

function applyFilters() {
    // Récupérer les valeurs des filtres
    state.filters.etb = document.getElementById('filterETB')?.value || '';
    state.filters.classification = document.getElementById('filterClassification')?.value || '';
    state.filters.contrat = document.getElementById('filterContrat')?.value || '';
    state.filters.search = document.getElementById('globalSearch')?.value.toLowerCase() || '';
    
    // Filtrer les données effectif
    state.filteredEffectif = state.data.effectif.filter(emp => {
        // Filtre ETB
        if (state.filters.etb && emp.etb !== state.filters.etb) return false;
        
        // Filtre Classification
        if (state.filters.classification) {
            const empClass = normalizeClassification(emp.classification);
            const filterClass = normalizeClassification(state.filters.classification);
            if (empClass !== filterClass) return false;
        }
        
        // Filtre Contrat
        if (state.filters.contrat && emp.contrat !== state.filters.contrat) return false;
        
        // Filtre Recherche
        if (state.filters.search) {
            const searchStr = `${emp.nom} ${emp.prenom} ${emp.matricule} ${emp.fonction}`.toLowerCase();
            if (!searchStr.includes(state.filters.search)) return false;
        }
        
        return true;
    });
    
    // Filtrer les données MS
    state.filteredMS = state.data.masseSalariale.filter(ms => {
        if (state.filters.etb) {
            const msEtb = ms.etb === 'Siége' ? 'Siège' : ms.etb;
            if (msEtb !== state.filters.etb) return false;
        }
        if (state.filters.classification) {
            const msClass = normalizeClassification(ms.classification);
            const filterClass = normalizeClassification(state.filters.classification);
            if (msClass !== filterClass) return false;
        }
        return true;
    });
    
    // Filtrer recrutements
    state.filteredRecrutements = state.data.recrutements.filter(r => {
        if (state.filters.etb && r.site !== state.filters.etb) return false;
        if (state.filters.classification) {
            const rClass = normalizeClassification(r.college);
            const filterClass = normalizeClassification(state.filters.classification);
            if (rClass !== filterClass) return false;
        }
        return true;
    });
    
    // Filtrer départs
    state.filteredDeparts = state.data.departs.filter(d => {
        if (state.filters.etb && d.site !== state.filters.etb) return false;
        if (state.filters.classification) {
            const dClass = normalizeClassification(d.college);
            const filterClass = normalizeClassification(state.filters.classification);
            if (dClass !== filterClass) return false;
        }
        return true;
    });
    
    // Réinitialiser la pagination
    state.currentPage = 1;
    
    // Mettre à jour l'affichage
    renderEffectifTable();
    renderMSTable();
    renderMouvementsTable();
    updateAllKPIs();
    updateCharts();
}

function resetFilters() {
    // Réinitialiser les valeurs
    document.getElementById('filterETB').value = '';
    document.getElementById('filterClassification').value = '';
    document.getElementById('filterContrat').value = '';
    document.getElementById('globalSearch').value = '';
    
    // Réinitialiser l'état
    state.filters = { etb: '', classification: '', contrat: '', search: '' };
    state.currentPage = 1;
    
    // Restaurer toutes les données
    state.filteredEffectif = [...state.data.effectif];
    state.filteredMS = [...state.data.masseSalariale];
    state.filteredRecrutements = [...state.data.recrutements];
    state.filteredDeparts = [...state.data.departs];
    
    // Mettre à jour l'affichage
    renderEffectifTable();
    renderMSTable();
    renderMouvementsTable();
    updateAllKPIs();
    updateCharts();
}

function normalizeClassification(value) {
    if (!value) return '';
    const normalized = value.toLowerCase().trim()
        .replace(/maitrise|maîtrise|matrise/gi, 'maitrise')
        .replace(/employé|employe|emploi/gi, 'employe')
        .replace(/cadre/gi, 'cadre');
    
    if (normalized.includes('cadre')) return 'cadre';
    if (normalized.includes('maitrise')) return 'maitrise';
    if (normalized.includes('employe')) return 'employe';
    return normalized;
}

// ============================================
// CALCUL DES KPIs
// ============================================

function updateAllKPIs() {
    const effectif = state.filteredEffectif;
    const ms = state.filteredMS;
    const recrutements = state.filteredRecrutements;
    const departs = state.filteredDeparts;
    
    const totalFTE = effectif.length;
    const hommes = effectif.filter(e => e.civilite === 'Mr').length;
    const femmes = totalFTE - hommes;
    const cdi = effectif.filter(e => e.contrat === 'CDI').length;
    const cio = totalFTE - cdi;
    
    // Calculs statistiques
    const avgAge = totalFTE > 0 ? Math.round(effectif.reduce((a, b) => a + (b.age || 0), 0) / totalFTE) : 0;
    const avgAnc = totalFTE > 0 ? Math.round(effectif.reduce((a, b) => a + (b.anciennete || 0), 0) / totalFTE) : 0;
    
    // Masse salariale
    const msJanvier = ms.reduce((a, b) => a + (b.janvier || 0), 0);
    const msYTD = msJanvier; // Pour janvier, YTD = Janvier
    
    // Indicateurs de turnover
    const nbrRecrutements = recrutements.length;
    const nbrDeparts = departs.length;
    const fteJanvierHorsRecrutements = totalFTE - nbrRecrutements;
    
    // Turnover = (Recrutements + Départs) / 2 / FTE janvier (hors recrutements)
    const turnover = fteJanvierHorsRecrutements > 0 
        ? ((nbrRecrutements + nbrDeparts) / 2 / fteJanvierHorsRecrutements * 100).toFixed(2) 
        : 0;
    
    // Taux de démission
    const nbrDemissions = departs.filter(d => d.motif && d.motif.toLowerCase().includes('démission')).length;
    const tauxDemission = fteJanvierHorsRecrutements > 0 
        ? (nbrDemissions / fteJanvierHorsRecrutements * 100).toFixed(2) 
        : 0;
    
    // Solde net
    const soldeNet = nbrRecrutements - nbrDeparts;
    
    // Taux de rétention
    const tauxRetention = totalFTE > 0 
        ? ((totalFTE - nbrDeparts) / (totalFTE - nbrRecrutements + nbrDeparts) * 100).toFixed(2) 
        : 0;
    
    // Taux de réussite période d'essai
    const echecsPE = departs.filter(d => d.motif && d.motif.toLowerCase().includes('essai')).length;
    const tauxReussitePE = nbrRecrutements > 0 
        ? ((nbrRecrutements - echecsPE) / nbrRecrutements * 100).toFixed(0) 
        : 100;
    
    // Taux d'accroissement
    const tauxAccroissement = fteJanvierHorsRecrutements > 0 
        ? (soldeNet / fteJanvierHorsRecrutements * 100).toFixed(2) 
        : 0;
    
    // Indicateurs additionnels
    // Taux d'encadrement (Cadres / Total)
    const nbrCadres = effectif.filter(e => normalizeClassification(e.classification) === 'cadre').length;
    const tauxEncadrement = totalFTE > 0 ? (nbrCadres / totalFTE * 100).toFixed(1) : 0;
    
    // Ratio H/F
    const ratioHF = femmes > 0 ? (hommes / femmes).toFixed(1) : hommes;
    
    // Taux de féminisation
    const tauxFeminisation = totalFTE > 0 ? (femmes / totalFTE * 100).toFixed(1) : 0;
    
    // Indice de vieillissement (>50 ans / total)
    const seniors = effectif.filter(e => (e.age || 0) >= 50).length;
    const indiceVieillissement = totalFTE > 0 ? (seniors / totalFTE * 100).toFixed(1) : 0;
    
    // Taux de stabilité (ancienneté > 5 ans / total)
    const stables = effectif.filter(e => (e.anciennete || 0) >= 5).length;
    const tauxStabilite = totalFTE > 0 ? (stables / totalFTE * 100).toFixed(1) : 0;
    
    // Coût moyen par collaborateur
    const coutMoyen = totalFTE > 0 ? (msJanvier / totalFTE) : 0;
    
    // MS par catégorie
    const msCadres = ms.filter(m => normalizeClassification(m.classification) === 'cadre')
        .reduce((a, b) => a + (b.janvier || 0), 0);
    const msMaitrise = ms.filter(m => normalizeClassification(m.classification) === 'maitrise')
        .reduce((a, b) => a + (b.janvier || 0), 0);
    const msEmployes = ms.filter(m => normalizeClassification(m.classification) === 'employe')
        .reduce((a, b) => a + (b.janvier || 0), 0);
    
    // Mettre à jour le DOM
    // Page Synthèse
    updateElement('kpi-fte', formatNumber(totalFTE));
    updateElement('kpi-gender', `${Math.round(hommes / totalFTE * 100)}%`);
    updateElement('kpi-age', `${avgAge} ans`);
    updateElement('kpi-anciennete', `${avgAnc} ans`);
    updateElement('kpi-ms-ytd', formatCurrency(msYTD));
    updateElement('kpi-recrutements', nbrRecrutements);
    updateElement('kpi-departs', nbrDeparts);
    updateElement('kpi-turnover', `${turnover}%`);
    
    // Nouveaux KPIs Synthèse
    updateElement('kpi-taux-encadrement', `${tauxEncadrement}%`);
    updateElement('kpi-taux-feminisation', `${tauxFeminisation}%`);
    updateElement('kpi-indice-vieillissement', `${indiceVieillissement}%`);
    updateElement('kpi-taux-stabilite', `${tauxStabilite}%`);
    
    // Page Effectifs
    updateElement('stat-total-fte', formatNumber(totalFTE));
    updateElement('stat-cdi', formatNumber(cdi));
    updateElement('stat-cio', formatNumber(cio));
    updateElement('stat-taux-cdi', `${(cdi / totalFTE * 100).toFixed(1)}%`);
    updateElement('stat-hommes', formatNumber(hommes));
    updateElement('stat-femmes', formatNumber(femmes));
    updateElement('stat-ratio-hf', ratioHF);
    updateElement('stat-age-moyen', `${avgAge} ans`);
    
    // Page Masse Salariale
    updateElement('ms-ytd', formatCurrency(msYTD));
    updateElement('ms-janvier', formatCurrency(msJanvier));
    updateElement('ms-per-capita', formatCurrency(coutMoyen));
    updateElement('ms-budget-annuel', formatCurrency(msJanvier * 12));
    updateElement('ms-cadres', formatCurrency(msCadres));
    updateElement('ms-maitrise', formatCurrency(msMaitrise));
    updateElement('ms-employes', formatCurrency(msEmployes));
    
    // Page Mouvements
    updateElement('ind-turnover', `${turnover}%`);
    updateElement('ind-demission', `${tauxDemission}%`);
    updateElement('ind-solde', soldeNet >= 0 ? `+${soldeNet}` : soldeNet);
    updateElement('ind-retention', `${tauxRetention}%`);
    updateElement('ind-essai', `${tauxReussitePE}%`);
    updateElement('ind-accroissement', `${tauxAccroissement}%`);
    updateElement('ind-recrutements', nbrRecrutements);
    updateElement('ind-departs', nbrDeparts);
    updateElement('ind-demissions', nbrDemissions);
    
    // Détails formules
    updateElement('detail-turnover', `(${nbrRecrutements} + ${nbrDeparts}) / 2 / ${fteJanvierHorsRecrutements} FTE`);
    updateElement('detail-demission', `${nbrDemissions} démissions / ${fteJanvierHorsRecrutements} FTE`);
    updateElement('detail-solde', `${nbrRecrutements} entrées - ${nbrDeparts} sorties`);
}

// ============================================
// GRAPHIQUES
// ============================================

function initCharts() {
    Chart.defaults.font.family = "'Plus Jakarta Sans', sans-serif";
    Chart.defaults.color = '#64748B';
    
    // Pyramide des âges
    createAgePyramidChart();
    
    // Pyramide d'ancienneté
    createAnciennetePyramidChart();
    
    // Répartition par classification
    createClassificationChart();
    
    // Répartition par ETB
    createETBChart();
    
    // Evolution MS
    createMSEvolutionChart();
    
    // Charts page Effectifs
    createFTEbyETBChart();
    createFTEbyClassificationChart();
    
    // Charts page MS
    createMSbyETBChart();
    createMSbyClassificationChart();
    
    // Charts page Mouvements
    createRecrutementsSiteChart();
    createMotifsDepartChart();
}

function createAgePyramidChart() {
    const ctx = document.getElementById('chartAgePyramid');
    if (!ctx) return;
    
    const ageData = calculateAgeDistribution();
    
    state.charts.agePyramid = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ageData.labels,
            datasets: [
                {
                    label: 'Hommes',
                    data: ageData.hommes,
                    backgroundColor: CONFIG.chartColors.orange,
                    borderRadius: 4
                },
                {
                    label: 'Femmes',
                    data: ageData.femmes,
                    backgroundColor: CONFIG.chartColors.teal,
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top' }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: CONFIG.chartColors.grayLight } },
                x: { grid: { display: false } }
            }
        }
    });
}

function createAnciennetePyramidChart() {
    const ctx = document.getElementById('chartAnciennetePyramid');
    if (!ctx) return;
    
    const ancData = calculateAncienneteDistribution();
    
    state.charts.anciennetePyramid = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ancData.labels,
            datasets: [{
                label: 'Effectif',
                data: ancData.values,
                backgroundColor: CONFIG.chartColors.blue,
                borderRadius: 4
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                x: { beginAtZero: true, grid: { color: CONFIG.chartColors.grayLight } },
                y: { grid: { display: false } }
            }
        }
    });
}

function createClassificationChart() {
    const ctx = document.getElementById('chartClassification');
    if (!ctx) return;
    
    const data = calculateClassificationDistribution();
    
    state.charts.classification = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Cadres', 'Maîtrise', 'Employés'],
            datasets: [{
                data: [data.cadres, data.maitrise, data.employes],
                backgroundColor: [CONFIG.chartColors.purple, CONFIG.chartColors.teal, CONFIG.chartColors.orange],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '65%',
            plugins: { legend: { position: 'bottom' } }
        }
    });
}

function createETBChart() {
    const ctx = document.getElementById('chartETB');
    if (!ctx) return;
    
    const data = calculateETBDistribution();
    
    state.charts.etb = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(data),
            datasets: [{
                data: Object.values(data),
                backgroundColor: [CONFIG.chartColors.orange, CONFIG.chartColors.blue, CONFIG.chartColors.teal],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '65%',
            plugins: { legend: { position: 'bottom' } }
        }
    });
}

function createMSEvolutionChart() {
    const ctx = document.getElementById('chartMSEvolution');
    if (!ctx) return;
    
    const moisLabels = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];
    const msData = calculateMSByMonth();
    const budget = msData[0] || 0;
    
    state.charts.msEvolution = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: moisLabels,
            datasets: [
                {
                    label: 'Réalisé',
                    data: msData,
                    backgroundColor: CONFIG.chartColors.orange,
                    borderRadius: 6
                },
                {
                    label: 'Budget',
                    data: Array(12).fill(budget),
                    backgroundColor: CONFIG.chartColors.grayLight,
                    borderRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top' },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.dataset.label}: ${formatCurrency(ctx.raw * 1000000)}`
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: CONFIG.chartColors.grayLight },
                    ticks: { callback: v => v.toFixed(1) + 'M' }
                },
                x: { grid: { display: false } }
            }
        }
    });
}

function createFTEbyETBChart() {
    const ctx = document.getElementById('chartFTEbyETB');
    if (!ctx) return;
    
    const data = calculateETBDistribution();
    
    state.charts.fteByETB = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(data),
            datasets: [{
                label: 'Effectif',
                data: Object.values(data),
                backgroundColor: [CONFIG.chartColors.orange, CONFIG.chartColors.blue, CONFIG.chartColors.teal],
                borderRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: CONFIG.chartColors.grayLight } },
                x: { grid: { display: false } }
            }
        }
    });
}

function createFTEbyClassificationChart() {
    const ctx = document.getElementById('chartFTEbyClassification');
    if (!ctx) return;
    
    const data = calculateClassificationDistribution();
    
    state.charts.fteByClassification = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Employés', 'Maîtrise', 'Cadres'],
            datasets: [{
                label: 'Effectif',
                data: [data.employes, data.maitrise, data.cadres],
                backgroundColor: [CONFIG.chartColors.orange, CONFIG.chartColors.teal, CONFIG.chartColors.purple],
                borderRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: CONFIG.chartColors.grayLight } },
                x: { grid: { display: false } }
            }
        }
    });
}

function createMSbyETBChart() {
    const ctx = document.getElementById('chartMSbyETB');
    if (!ctx) return;
    
    const data = calculateMSbyETB();
    
    state.charts.msByETB = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: Object.keys(data),
            datasets: [{
                data: Object.values(data).map(v => v / 1000000),
                backgroundColor: [CONFIG.chartColors.orange, CONFIG.chartColors.blue, CONFIG.chartColors.teal],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.label}: ${ctx.raw.toFixed(2)}M MAD`
                    }
                }
            }
        }
    });
}

function createMSbyClassificationChart() {
    const ctx = document.getElementById('chartMSbyClassification');
    if (!ctx) return;
    
    const data = calculateMSbyClassification();
    
    state.charts.msByClassification = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: Object.keys(data),
            datasets: [{
                data: Object.values(data).map(v => v / 1000000),
                backgroundColor: [CONFIG.chartColors.purple, CONFIG.chartColors.teal, CONFIG.chartColors.orange],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.label}: ${ctx.raw.toFixed(2)}M MAD`
                    }
                }
            }
        }
    });
}

function createRecrutementsSiteChart() {
    const ctx = document.getElementById('chartRecrutementsSite');
    if (!ctx) return;
    
    const data = {};
    state.filteredRecrutements.forEach(r => {
        data[r.site] = (data[r.site] || 0) + 1;
    });
    
    state.charts.recrutementsSite = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(data),
            datasets: [{
                label: 'Recrutements',
                data: Object.values(data),
                backgroundColor: CONFIG.chartColors.success,
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: CONFIG.chartColors.grayLight } },
                x: { grid: { display: false } }
            }
        }
    });
}

function createMotifsDepartChart() {
    const ctx = document.getElementById('chartMotifsDepart');
    if (!ctx) return;
    
    const data = {};
    state.filteredDeparts.forEach(d => {
        const motif = d.motif || 'Autre';
        data[motif] = (data[motif] || 0) + 1;
    });
    
    state.charts.motifsDepart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(data),
            datasets: [{
                data: Object.values(data),
                backgroundColor: [CONFIG.chartColors.danger, CONFIG.chartColors.gold, CONFIG.chartColors.gray],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '60%',
            plugins: { legend: { position: 'bottom' } }
        }
    });
}

function updateCharts() {
    // Mettre à jour tous les graphiques
    if (state.charts.agePyramid) {
        const ageData = calculateAgeDistribution();
        state.charts.agePyramid.data.datasets[0].data = ageData.hommes;
        state.charts.agePyramid.data.datasets[1].data = ageData.femmes;
        state.charts.agePyramid.update();
    }
    
    if (state.charts.anciennetePyramid) {
        const ancData = calculateAncienneteDistribution();
        state.charts.anciennetePyramid.data.datasets[0].data = ancData.values;
        state.charts.anciennetePyramid.update();
    }
    
    if (state.charts.classification) {
        const data = calculateClassificationDistribution();
        state.charts.classification.data.datasets[0].data = [data.cadres, data.maitrise, data.employes];
        state.charts.classification.update();
    }
    
    if (state.charts.etb) {
        const data = calculateETBDistribution();
        state.charts.etb.data.labels = Object.keys(data);
        state.charts.etb.data.datasets[0].data = Object.values(data);
        state.charts.etb.update();
    }
    
    if (state.charts.fteByETB) {
        const data = calculateETBDistribution();
        state.charts.fteByETB.data.labels = Object.keys(data);
        state.charts.fteByETB.data.datasets[0].data = Object.values(data);
        state.charts.fteByETB.update();
    }
    
    if (state.charts.fteByClassification) {
        const data = calculateClassificationDistribution();
        state.charts.fteByClassification.data.datasets[0].data = [data.employes, data.maitrise, data.cadres];
        state.charts.fteByClassification.update();
    }
    
    // Mettre à jour les graphiques de mouvements
    if (state.charts.recrutementsSite) {
        const data = {};
        state.filteredRecrutements.forEach(r => {
            data[r.site] = (data[r.site] || 0) + 1;
        });
        state.charts.recrutementsSite.data.labels = Object.keys(data);
        state.charts.recrutementsSite.data.datasets[0].data = Object.values(data);
        state.charts.recrutementsSite.update();
    }
    
    if (state.charts.motifsDepart) {
        const data = {};
        state.filteredDeparts.forEach(d => {
            const motif = d.motif || 'Autre';
            data[motif] = (data[motif] || 0) + 1;
        });
        state.charts.motifsDepart.data.labels = Object.keys(data);
        state.charts.motifsDepart.data.datasets[0].data = Object.values(data);
        state.charts.motifsDepart.update();
    }
}

// ============================================
// CALCULS DE DONNÉES
// ============================================

function calculateAgeDistribution() {
    const bins = [
        { label: '<25', min: 0, max: 25 },
        { label: '25-30', min: 25, max: 30 },
        { label: '30-35', min: 30, max: 35 },
        { label: '35-40', min: 35, max: 40 },
        { label: '40-45', min: 40, max: 45 },
        { label: '45-50', min: 45, max: 50 },
        { label: '50-55', min: 50, max: 55 },
        { label: '55-60', min: 55, max: 60 },
        { label: '>60', min: 60, max: 100 }
    ];
    
    const hommes = bins.map(bin => 
        state.filteredEffectif.filter(e => 
            e.civilite === 'Mr' && (e.age || 0) >= bin.min && (e.age || 0) < bin.max
        ).length
    );
    
    const femmes = bins.map(bin => 
        state.filteredEffectif.filter(e => 
            e.civilite !== 'Mr' && (e.age || 0) >= bin.min && (e.age || 0) < bin.max
        ).length
    );
    
    return {
        labels: bins.map(b => b.label),
        hommes,
        femmes
    };
}

function calculateAncienneteDistribution() {
    const bins = [
        { label: '<1 an', min: 0, max: 1 },
        { label: '1-3 ans', min: 1, max: 3 },
        { label: '3-5 ans', min: 3, max: 5 },
        { label: '5-10 ans', min: 5, max: 10 },
        { label: '10-15 ans', min: 10, max: 15 },
        { label: '15-20 ans', min: 15, max: 20 },
        { label: '20-25 ans', min: 20, max: 25 },
        { label: '25-30 ans', min: 25, max: 30 },
        { label: '>30 ans', min: 30, max: 100 }
    ];
    
    const values = bins.map(bin => 
        state.filteredEffectif.filter(e => 
            (e.anciennete || 0) >= bin.min && (e.anciennete || 0) < bin.max
        ).length
    );
    
    return {
        labels: bins.map(b => b.label),
        values
    };
}

function calculateClassificationDistribution() {
    return {
        cadres: state.filteredEffectif.filter(e => normalizeClassification(e.classification) === 'cadre').length,
        maitrise: state.filteredEffectif.filter(e => normalizeClassification(e.classification) === 'maitrise').length,
        employes: state.filteredEffectif.filter(e => normalizeClassification(e.classification) === 'employe').length
    };
}

function calculateETBDistribution() {
    const data = {};
    state.filteredEffectif.forEach(e => {
        data[e.etb] = (data[e.etb] || 0) + 1;
    });
    return data;
}

function calculateMSByMonth() {
    const months = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'];
    return months.map(m => {
        const total = state.filteredMS.reduce((a, b) => a + (b[m] || 0), 0);
        return total / 1000000;
    });
}

function calculateMSbyETB() {
    const data = {};
    state.filteredMS.forEach(ms => {
        const etb = ms.etb === 'Siége' ? 'Siège' : ms.etb;
        data[etb] = (data[etb] || 0) + (ms.janvier || 0);
    });
    return data;
}

function calculateMSbyClassification() {
    const data = { 'Cadres': 0, 'Maîtrise': 0, 'Employés': 0 };
    state.filteredMS.forEach(ms => {
        const cls = normalizeClassification(ms.classification);
        if (cls === 'cadre') data['Cadres'] += (ms.janvier || 0);
        else if (cls === 'maitrise') data['Maîtrise'] += (ms.janvier || 0);
        else data['Employés'] += (ms.janvier || 0);
    });
    return data;
}

// ============================================
// TABLEAUX
// ============================================

function renderEffectifTable() {
    const tbody = document.getElementById('effectifTableBody');
    if (!tbody) return;
    
    const start = (state.currentPage - 1) * CONFIG.itemsPerPage;
    const end = start + CONFIG.itemsPerPage;
    const pageData = state.filteredEffectif.slice(start, end);
    
    tbody.innerHTML = pageData.map(emp => `
        <tr>
            <td><strong>${emp.matricule}</strong></td>
            <td>${emp.nom || ''}</td>
            <td>${emp.prenom || ''}</td>
            <td>${emp.etb || ''}</td>
            <td><span class="badge badge-${getBadgeClass(emp.classification)}">${emp.classification || ''}</span></td>
            <td>${emp.fonction || ''}</td>
            <td>${emp.age || 0} ans</td>
            <td>${emp.anciennete || 0} ans</td>
            <td><span class="badge ${emp.contrat === 'CDI' ? 'badge-success' : 'badge-warning'}">${emp.contrat || ''}</span></td>
        </tr>
    `).join('');
    
    renderPagination();
}

function renderPagination() {
    const total = state.filteredEffectif.length;
    const totalPages = Math.ceil(total / CONFIG.itemsPerPage);
    const start = (state.currentPage - 1) * CONFIG.itemsPerPage + 1;
    const end = Math.min(state.currentPage * CONFIG.itemsPerPage, total);
    
    const infoEl = document.getElementById('paginationInfo');
    if (infoEl) {
        infoEl.textContent = `Affichage ${start}-${end} sur ${total} collaborateurs`;
    }
    
    const buttonsEl = document.getElementById('paginationButtons');
    if (!buttonsEl) return;
    
    let html = `<button class="pagination-btn" onclick="changePage(${state.currentPage - 1})" ${state.currentPage === 1 ? 'disabled' : ''}><i class="fas fa-chevron-left"></i></button>`;
    
    const maxVisible = 5;
    let startPage = Math.max(1, state.currentPage - Math.floor(maxVisible / 2));
    let endPage = Math.min(totalPages, startPage + maxVisible - 1);
    
    if (endPage - startPage + 1 < maxVisible) {
        startPage = Math.max(1, endPage - maxVisible + 1);
    }
    
    if (startPage > 1) {
        html += `<button class="pagination-btn" onclick="changePage(1)">1</button>`;
        if (startPage > 2) html += `<span style="padding: 0 8px; color: var(--text-muted);">...</span>`;
    }
    
    for (let i = startPage; i <= endPage; i++) {
        html += `<button class="pagination-btn ${i === state.currentPage ? 'active' : ''}" onclick="changePage(${i})">${i}</button>`;
    }
    
    if (endPage < totalPages) {
        if (endPage < totalPages - 1) html += `<span style="padding: 0 8px; color: var(--text-muted);">...</span>`;
        html += `<button class="pagination-btn" onclick="changePage(${totalPages})">${totalPages}</button>`;
    }
    
    html += `<button class="pagination-btn" onclick="changePage(${state.currentPage + 1})" ${state.currentPage === totalPages ? 'disabled' : ''}><i class="fas fa-chevron-right"></i></button>`;
    
    buttonsEl.innerHTML = html;
}

function changePage(page) {
    const totalPages = Math.ceil(state.filteredEffectif.length / CONFIG.itemsPerPage);
    if (page < 1 || page > totalPages) return;
    state.currentPage = page;
    renderEffectifTable();
}

function renderMSTable() {
    const tbody = document.getElementById('msTableBody');
    if (!tbody) return;
    
    const msData = state.filteredMS.slice(0, 50);
    
    tbody.innerHTML = msData.map(item => `
        <tr>
            <td>${item.intitule || ''}</td>
            <td>${item.etb || ''}</td>
            <td><span class="badge badge-${getBadgeClass(item.classification)}">${item.classification || ''}</span></td>
            <td>${formatNumber(item.janvier || 0)} MAD</td>
            <td>${formatNumber(item.janvier || 0)} MAD</td>
        </tr>
    `).join('');
}

function renderMouvementsTable() {
    // Recrutements
    const recBody = document.getElementById('recrutementsTableBody');
    if (recBody) {
        recBody.innerHTML = state.filteredRecrutements.map(r => `
            <tr>
                <td><strong>${r.matricule}</strong></td>
                <td>${r.nom} ${r.prenom}</td>
                <td>${r.site}</td>
                <td><span class="badge badge-${getBadgeClass(r.college)}">${r.college}</span></td>
                <td>${formatDate(r.date)}</td>
                <td>${r.fonction}</td>
            </tr>
        `).join('');
    }
    
    // Départs
    const depBody = document.getElementById('departsTableBody');
    if (depBody) {
        depBody.innerHTML = state.filteredDeparts.map(d => `
            <tr>
                <td><strong>${d.matricule}</strong></td>
                <td>${d.nom} ${d.prenom}</td>
                <td>${d.site}</td>
                <td><span class="badge badge-${getBadgeClass(d.college)}">${d.college}</span></td>
                <td>${formatDate(d.date)}</td>
                <td><span class="badge ${getMotifBadge(d.motif)}">${d.motif}</span></td>
            </tr>
        `).join('');
    }
}

// ============================================
// UTILITAIRES
// ============================================

function formatNumber(num) {
    if (num === null || num === undefined) return '0';
    return new Intl.NumberFormat('fr-FR').format(Math.round(num));
}

function formatCurrency(num) {
    if (num === null || num === undefined) return '0 MAD';
    if (num >= 1000000) {
        return (num / 1000000).toFixed(1) + 'M';
    }
    if (num >= 1000) {
        return (num / 1000).toFixed(0) + 'K';
    }
    return formatNumber(num);
}

function formatDate(dateStr) {
    if (!dateStr) return '-';
    try {
        return new Date(dateStr).toLocaleDateString('fr-FR');
    } catch {
        return dateStr;
    }
}

function formatDateTime(dateStr) {
    if (!dateStr) return '-';
    try {
        return new Date(dateStr).toLocaleString('fr-FR');
    } catch {
        return dateStr;
    }
}

function getBadgeClass(classification) {
    const cls = normalizeClassification(classification);
    if (cls === 'cadre') return 'cadre';
    if (cls === 'maitrise') return 'maitrise';
    return 'employe';
}

function getMotifBadge(motif) {
    if (!motif) return 'badge-info';
    const m = motif.toLowerCase();
    if (m.includes('démission')) return 'badge-danger';
    if (m.includes('essai')) return 'badge-warning';
    if (m.includes('décès')) return 'badge-info';
    return 'badge-info';
}

function updateElement(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

function showError(message) {
    alert(message);
}

async function refreshData() {
    const btn = document.getElementById('btnRefresh');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i>';
    }
    
    await loadData();
    initializeApp();
    
    if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-sync-alt"></i>';
    }
}

// ============================================
// EXPORT
// ============================================

function exportEffectifCSV() {
    let csv = '\uFEFF'; // BOM for Excel
    csv += 'Matricule;Nom;Prénom;Établissement;Classification;Fonction;Âge;Ancienneté;Contrat\n';
    
    state.filteredEffectif.forEach(emp => {
        csv += `${emp.matricule};${emp.nom};${emp.prenom};${emp.etb};${emp.classification};${emp.fonction};${emp.age};${emp.anciennete};${emp.contrat}\n`;
    });
    
    downloadFile(csv, 'effectif_export.csv', 'text/csv;charset=utf-8');
}

function exportMSCSV() {
    let csv = '\uFEFF';
    csv += 'Rubrique;Établissement;Classification;Janvier;YTD\n';
    
    state.filteredMS.forEach(item => {
        csv += `${item.intitule};${item.etb};${item.classification};${item.janvier};${item.janvier}\n`;
    });
    
    downloadFile(csv, 'masse_salariale_export.csv', 'text/csv;charset=utf-8');
}

function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Exposer les fonctions globalement
window.showPage = showPage;
window.applyFilters = applyFilters;
window.resetFilters = resetFilters;
window.changePage = changePage;
window.refreshData = refreshData;
window.exportEffectifCSV = exportEffectifCSV;
window.exportMSCSV = exportMSCSV;
