import { useState, useMemo } from 'react'
import { 
  Search, 
  ChevronDown, 
  ChevronUp, 
  X, 
  RotateCcw, 
  FileSpreadsheet, 
  GraduationCap, 
  Info, 
  Sliders, 
  CheckCircle2, 
  AlertCircle, 
  XCircle, 
  Calendar, 
  Hash, 
  User 
} from 'lucide-react'
import rawStudentsData from './data/students.json'
import { exportCohortToExcel as exportCohort, exportIndividualToExcel as exportIndividual } from './utils/excelExport'


// Standard specialties labels mapping for displays
const SPEC_SHORT_LABELS = {
  "Mathématiques": "Maths",
  "Physique-chimie": "Phys-Chimie",
  "Sciences de la vie et de la terre": "SVT",
  "Sciences économiques et sociales": "SES",
  "Histoire-géographie, géopolitique et sciences politiques": "HGGSP",
  "Humanités, littérature et philosophie": "HLP",
  "Langues, littératures et cultures étrangères et régionales - Anglais, monde contemporain": "LLCE"
}

function App() {
  // 1. Core States
  const [students, setStudents] = useState(rawStudentsData)
  
  // Simulations state: stores custom simulated grades for each student
  // Initialized with 10 for terminal exams and simulated EPS, and French grades from data (or 10 if missing)
  const [simulations, setSimulations] = useState(() => {
    const initSims = {}
    rawStudentsData.forEach(student => {
      initSims[student.full_id] = {
        spe1: 10.0,
        spe2: 10.0,
        philo: 10.0,
        grandOral: 10.0,
        eps: (student.eps_grade !== undefined && student.eps_grade !== 'dispense') ? student.eps_grade : 10.0,
        frenchWritten: student.french_written !== null ? student.french_written : 10.0,
        frenchOral: student.french_oral !== null ? student.french_oral : 10.0,
        isFrenchMissing: student.french_written === null
      }
    })
    return initSims
  })

  // Selected student for simulation sheet (modal / side drawer)
  const [selectedStudentId, setSelectedStudentId] = useState(null)
  
  // Accordion state for Baccalauréat rules
  const [isRulesExpanded, setIsRulesExpanded] = useState(false)
  
  // Filters state
  const [searchQuery, setSearchQuery] = useState('')
  const [filterSpec, setFilterSpec] = useState('')
  const [filterStatus, setFilterStatus] = useState('')
  const [filterMention, setFilterMention] = useState('')
  
  // Sorting state for table
  const [sortField, setSortField] = useState('name')
  const [sortDirection, setSortDirection] = useState('asc')

  // 2. Calculations Helper Functions
  const calculateStudentResults = (student, sim) => {
    // A. Continuous Control Points (40 Coefficients normally, 34 if EPS is dispensed)
    const lp001 = student.grades['LP001']?.value ?? 10.0 // HG 1ère
    const lt001 = student.grades['LT001']?.value ?? 10.0 // HG Term
    const hgPoints = (lp001 * 3) + (lt001 * 3) // Coef 6
    
    const lp002 = student.grades['LP002']?.value ?? 10.0 // EMC 1ère
    const lt002 = student.grades['LT002']?.value ?? 10.0 // EMC Term
    const emcPoints = (lp002 * 1) + (lt002 * 1) // Coef 2
    
    const lp003 = student.grades['LP003']?.value ?? 10.0 // LVA 1ère
    const lt003 = student.grades['LT003']?.value ?? 10.0 // LVA Term
    const lvaPoints = (lp003 * 3) + (lt003 * 3) // Coef 6
    
    const lp004 = student.grades['LP004']?.value ?? 10.0 // LVB 1ère
    const lt004 = student.grades['LT004']?.value ?? 10.0 // LVB Term
    const lvbPoints = (lp004 * 3) + (lt004 * 3) // Coef 6
    
    const ensSci1 = student.grades['LP005']?.value ?? student.grades['LP007']?.value ?? 10.0 // Sci 1ère
    const lt005 = student.grades['LT005']?.value ?? 10.0 // Sci Term
    const sciPoints = (ensSci1 * 3) + (lt005 * 3) // Coef 6
    
    const droppedSpec = student.dropped_specialty?.value ?? 10.0 // Specialty dropped 1ère (Coef 8)
    const droppedSpecPoints = droppedSpec * 8
    
    // EPS logic:
    let epsPoints = 0
    let ccCoef = 40.0
    let isEpsDispensed = false
    
    if (student.eps_grade === 'dispense') {
      isEpsDispensed = true
      ccCoef = 34.0 // 40 - 6
    } else if (student.eps_grade !== undefined) {
      epsPoints = student.eps_grade * 6
    } else {
      epsPoints = sim.eps * 6
    }
    
    const totalCcPoints = hgPoints + emcPoints + lvaPoints + lvbPoints + sciPoints + droppedSpecPoints + epsPoints
    const ccAverage = totalCcPoints / ccCoef
    
    // B. Terminal Exams Points (60 Coefficients)
    const frWrittenPoints = sim.frenchWritten * 5
    const frOralPoints = sim.frenchOral * 5
    const spe1Points = sim.spe1 * 16
    const spe2Points = sim.spe2 * 16
    const philoPoints = sim.philo * 8
    const grandOralPoints = sim.grandOral * 10
    
    const totalTerminalPoints = frWrittenPoints + frOralPoints + spe1Points + spe2Points + philoPoints + grandOralPoints
    const terminalAverage = totalTerminalPoints / 60.0
    
    // C. Overall Weighted Average (100 Coefficients or 94 if EPS is dispensed)
    const totalBacCoef = isEpsDispensed ? 94.0 : 100.0
    const finalAverage = (totalCcPoints + totalTerminalPoints) / totalBacCoef
    
    // D. Status and Mentions
    let status = 'Refusé'
    let mention = 'Aucune'
    
    if (finalAverage >= 10.0) {
      status = 'Admis'
      if (finalAverage >= 18.0) {
        mention = 'Très bien avec félicitations'
      } else if (finalAverage >= 16.0) {
        mention = 'Très bien'
      } else if (finalAverage >= 14.0) {
        mention = 'Bien'
      } else if (finalAverage >= 12.0) {
        mention = 'Assez bien'
      } else {
        mention = 'Admis sans mention'
      }
    } else if (finalAverage >= 8.0) {
      status = 'Rattrapage'
      mention = 'Admis au second groupe'
    }
    
    return {
      ccAverage,
      terminalAverage,
      finalAverage,
      status,
      mention,
      totalCcPoints,
      totalTerminalPoints,
      isEpsDispensed
    }
  }

  // 3. Computed cohort data based on active simulations
  const cohortData = useMemo(() => {
    const computed = students.map(s => {
      const sim = simulations[s.full_id]
      const results = calculateStudentResults(s, sim)
      return {
        ...s,
        results
      }
    })
    
    const totalCount = computed.length
    const admittedCount = computed.filter(s => s.results.status === 'Admis').length
    const admittedPercent = ((admittedCount / totalCount) * 100).toFixed(0)
    
    const cohortAverage = computed.reduce((acc, curr) => acc + curr.results.finalAverage, 0) / totalCount
    
    const mentionsCounts = {
      felicitations: computed.filter(s => s.results.mention === 'Très bien avec félicitations').length,
      tb: computed.filter(s => s.results.mention === 'Très bien').length,
      b: computed.filter(s => s.results.mention === 'Bien').length,
      ab: computed.filter(s => s.results.mention === 'Assez bien').length,
      sans: computed.filter(s => s.results.mention === 'Admis sans mention').length,
      rattrapage: computed.filter(s => s.results.status === 'Rattrapage').length,
      refuse: computed.filter(s => s.results.status === 'Refusé').length
    }
    
    const totalMentions = mentionsCounts.felicitations + mentionsCounts.tb + mentionsCounts.b + mentionsCounts.ab
    
    return {
      computedStudents: computed,
      totalCount,
      admittedCount,
      admittedPercent,
      cohortAverage,
      mentionsCounts,
      totalMentions
    }
  }, [students, simulations])

  // 4. Handle input updates for student simulations
  const handleGradeChange = (studentId, field, val) => {
    let numericVal = parseFloat(val)
    if (isNaN(numericVal)) numericVal = 0.0
    if (numericVal < 0) numericVal = 0.0
    if (numericVal > 20) numericVal = 20.0
    
    setSimulations(prev => ({
      ...prev,
      [studentId]: {
        ...prev[studentId],
        [field]: numericVal
      }
    }))
  }

  const resetStudentSimulation = (studentId) => {
    const student = students.find(s => s.full_id === studentId)
    if (!student) return
    
    setSimulations(prev => ({
      ...prev,
      [studentId]: {
        spe1: 10.0,
        spe2: 10.0,
        philo: 10.0,
        grandOral: 10.0,
        eps: (student.eps_grade !== undefined && student.eps_grade !== 'dispense') ? student.eps_grade : 10.0,
        frenchWritten: student.french_written !== null ? student.french_written : 10.0,
        frenchOral: student.french_oral !== null ? student.french_oral : 10.0,
        isFrenchMissing: student.french_written === null
      }
    }))
  }

  const resetAllSimulations = () => {
    const newSims = {}
    students.forEach(student => {
      newSims[student.full_id] = {
        spe1: 10.0,
        spe2: 10.0,
        philo: 10.0,
        grandOral: 10.0,
        eps: (student.eps_grade !== undefined && student.eps_grade !== 'dispense') ? student.eps_grade : 10.0,
        frenchWritten: student.french_written !== null ? student.french_written : 10.0,
        frenchOral: student.french_oral !== null ? student.french_oral : 10.0,
        isFrenchMissing: student.french_written === null
      }
    })
    setSimulations(newSims)
  }

  // 5. Excel Export Functions
  const exportCohortToExcel = () => {
    exportCohort(cohortData.computedStudents, simulations)
  }

  const exportIndividualToExcel = (student, sim, results) => {
    exportIndividual(student, sim, results)
  }

  // 6. Filter & Sort Candidates List
  const filteredStudents = useMemo(() => {
    return cohortData.computedStudents.filter(student => {
      // Name search
      const nameMatch = student.name.toLowerCase().includes(searchQuery.toLowerCase())
      
      // Specialty filter
      const specMatch = !filterSpec || student.specialties.includes(filterSpec)
      
      // Status filter
      const statusMatch = !filterStatus || student.results.status === filterStatus
      
      // Mention filter
      const mentionMatch = !filterMention || student.results.mention === filterMention
      
      return nameMatch && specMatch && statusMatch && mentionMatch
    }).sort((a, b) => {
      let valA, valB
      if (sortField === 'name') {
        valA = a.family_name.toLowerCase()
        valB = b.family_name.toLowerCase()
      } else if (sortField === 'cc') {
        valA = a.results.ccAverage
        valB = b.results.ccAverage
      } else if (sortField === 'terminal') {
        valA = a.results.terminalAverage
        valB = b.results.terminalAverage
      } else if (sortField === 'final') {
        valA = a.results.finalAverage
        valB = b.results.finalAverage
      }
      
      if (valA < valB) return sortDirection === 'asc' ? -1 : 1
      if (valA > valB) return sortDirection === 'asc' ? 1 : -1
      return 0
    })
  }, [cohortData, searchQuery, filterSpec, filterStatus, filterMention, sortField, sortDirection])

  const toggleSort = (field) => {
    if (sortField === field) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')
    } else {
      setSortField(field)
      setSortDirection('asc')
    }
  }

  // Find currently selected student details
  const activeStudent = useMemo(() => {
    if (!selectedStudentId) return null
    return cohortData.computedStudents.find(s => s.full_id === selectedStudentId)
  }, [selectedStudentId, cohortData])

  const activeSim = useMemo(() => {
    if (!selectedStudentId) return null
    return simulations[selectedStudentId]
  }, [selectedStudentId, simulations])

  // 7. Calculate "Moyenne terminale minimale" for active student
  const minRequiredAverages = useMemo(() => {
    if (!activeStudent || !activeSim) return []
    
    // Points already acquired:
    // Total CC:
    const lp001 = activeStudent.grades['LP001']?.value ?? 10.0
    const lt001 = activeStudent.grades['LT001']?.value ?? 10.0
    const hgPoints = (lp001 * 3) + (lt001 * 3)
    
    const lp002 = activeStudent.grades['LP002']?.value ?? 10.0
    const lt002 = activeStudent.grades['LT002']?.value ?? 10.0
    const emcPoints = (lp002 * 1) + (lt002 * 1)
    
    const lp003 = activeStudent.grades['LP003']?.value ?? 10.0
    const lt003 = activeStudent.grades['LT003']?.value ?? 10.0
    const lvaPoints = (lp003 * 3) + (lt003 * 3)
    
    const lp004 = activeStudent.grades['LP004']?.value ?? 10.0
    const lt004 = activeStudent.grades['LT004']?.value ?? 10.0
    const lvbPoints = (lp004 * 3) + (lt004 * 3)
    
    const ensSci1 = activeStudent.grades['LP005']?.value ?? activeStudent.grades['LP007']?.value ?? 10.0
    const lt005 = activeStudent.grades['LT005']?.value ?? 10.0
    const sciPoints = (ensSci1 * 3) + (lt005 * 3)
    
    const droppedSpec = activeStudent.dropped_specialty?.value ?? 10.0
    const droppedSpecPoints = droppedSpec * 8
    
    // EPS dynamic check
    let epsPoints = 0
    let isEpsDispensed = false
    
    if (activeStudent.eps_grade === 'dispense') {
      isEpsDispensed = true;
    } else if (activeStudent.eps_grade !== undefined) {
      epsPoints = activeStudent.eps_grade * 6;
    } else {
      epsPoints = activeSim.eps * 6;
    }
    
    const totalCcPoints = hgPoints + emcPoints + lvaPoints + lvbPoints + sciPoints + droppedSpecPoints + epsPoints
    
    // French: written (coef 5) + oral (coef 5)
    const frenchPoints = (activeSim.frenchWritten * 5) + (activeSim.frenchOral * 5)
    
    const acquiredPoints = totalCcPoints + frenchPoints
    const remainingCoef = 50 // Spé 1 (16) + Spé 2 (16) + Philo (8) + Grand Oral (10)
    const totalBacCoef = isEpsDispensed ? 94.0 : 100.0
    
    const targets = [
      { label: "Baccalauréat (Admis)", target: 10.0 },
      { label: "Mention Assez Bien", target: 12.0 },
      { label: "Mention Bien", target: 14.0 },
      { label: "Mention Très Bien", target: 16.0 },
      { label: "Mention Très Bien avec Félicitations", target: 18.0 }
    ]
    
    return targets.map(t => {
      const neededPoints = (t.target * totalBacCoef) - acquiredPoints
      if (neededPoints <= 0) {
        return { label: t.label, target: t.target, note: "Déjà acquise", status: "acquired" }
      }
      
      const noteMin = neededPoints / remainingCoef
      if (noteMin > 20.0) {
        return { label: t.label, target: t.target, note: "Impossible (> 20)", status: "impossible" }
      }
      
      return { 
        label: t.label, 
        target: t.target, 
        note: noteMin.toFixed(2).replace('.', ',') + " / 20", 
        status: "normal" 
      }
    })
  }, [activeStudent, activeSim])

  return (
    <div className="container">
      {/* HEADER SECTION */}
      <header className="header">
        <div className="header-meta">Session Juin 2026 • Lycée Français Jacques Prévert (Saly)</div>
        <div className="header-flex">
          <div className="title-area">
            <h1 className="main-title">Simulateur de résultats</h1>
            <p className="subtitle">
              Projection à partir du contrôle continu réel et des notes d'épreuves anticipées de français enregistrées. Les épreuves terminales de Terminale sont calibrées à 10/20 par défaut et peuvent être ajustées par candidat.
            </p>
          </div>
          <div className="action-area">
            <button className="btn" onClick={resetAllSimulations} title="Réinitialiser toutes les simulations à 10/20">
              <RotateCcw size={16} />
              Réinitialiser cohorte
            </button>
            <button className="btn btn-primary" onClick={exportCohortToExcel}>
              <FileSpreadsheet size={16} />
              Exporter Excel
            </button>
          </div>
        </div>
      </header>

      {/* COHORT SUMMARY METRICS */}
      <section className="summary-grid">
        <div className="summary-card">
          <div className="summary-label">Candidats</div>
          <div className="summary-val-area">
            <div className="summary-value">{cohortData.totalCount}</div>
            <div className="summary-subtext">élèves</div>
          </div>
        </div>
        
        <div className="summary-card">
          <div className="summary-label">Admis projetés</div>
          <div className="summary-val-area">
            <div className="summary-value">{cohortData.admittedCount}</div>
            <div className="summary-subtext">/{cohortData.totalCount}</div>
            <span className="summary-badge badge-admis">{cohortData.admittedPercent}%</span>
          </div>
        </div>

        <div className="summary-card">
          <div className="summary-label">Moyenne cohorte</div>
          <div className="summary-val-area">
            <div className="summary-value">{cohortData.cohortAverage.toFixed(2).replace('.', ',')}</div>
            <div className="summary-subtext">/20</div>
          </div>
        </div>

        <div className="summary-card">
          <div className="summary-label">Mentions</div>
          <div className="summary-val-area">
            <div className="summary-value">{cohortData.totalMentions}</div>
            <div className="summary-subtext">projetées</div>
          </div>
        </div>
      </section>

      {/* DETAILED MENTIONS COUNTER */}
      <div className="mentions-sub-bar">
        <div className="mentions-sub-item">
          <span className="mentions-sub-count">{cohortData.mentionsCounts.felicitations}</span> Très Bien (Fél.)
        </div>
        <div className="mentions-sub-item">
          <span className="mentions-sub-count">{cohortData.mentionsCounts.tb}</span> Très Bien
        </div>
        <div className="mentions-sub-item">
          <span className="mentions-sub-count">{cohortData.mentionsCounts.b}</span> Bien
        </div>
        <div className="mentions-sub-item">
          <span className="mentions-sub-count">{cohortData.mentionsCounts.ab}</span> Assez Bien
        </div>
        <div className="mentions-sub-item">
          <span className="mentions-sub-count">{cohortData.mentionsCounts.sans}</span> Sans Mention
        </div>
        <div className="mentions-sub-item">
          <span className="mentions-sub-count">{cohortData.mentionsCounts.rattrapage}</span> Rattrapage
        </div>
        <div className="mentions-sub-item">
          <span className="mentions-sub-count">{cohortData.mentionsCounts.refuse}</span> Refusé
        </div>
      </div>

      {/* COEFFICIENTS & RULES ACCORDION */}
      <div className="accordion">
        <button 
          className="accordion-trigger" 
          onClick={() => setIsRulesExpanded(!isRulesExpanded)}
        >
          <span>Règles de calcul et coefficients appliqués (Bac 2026)</span>
          {isRulesExpanded ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
        </button>
        {isRulesExpanded && (
          <div className="accordion-content">
            <div className="rule-column">
              <h4>Contrôle Continu (40%)</h4>
              <ul className="rule-list">
                <li className="rule-item"><span>Histoire-Géographie (1ère + Term)</span><span className="rule-coef">Coef. 6</span></li>
                <li className="rule-item"><span>Langue Vivante A (1ère + Term)</span><span className="rule-coef">Coef. 6</span></li>
                <li className="rule-item"><span>Langue Vivante B (1ère + Term)</span><span className="rule-coef">Coef. 6</span></li>
                <li className="rule-item"><span>Enseignement Scientifique (1ère + Term)</span><span className="rule-coef">Coef. 6</span></li>
                <li className="rule-item"><span>EPS (Terminale uniquement - simulé)</span><span className="rule-coef">Coef. 6</span></li>
                <li className="rule-item"><span>EMC (1ère + Term)</span><span className="rule-coef">Coef. 2</span></li>
                <li className="rule-item"><span>Spécialité abandonnée en 1ère</span><span className="rule-coef">Coef. 8</span></li>
              </ul>
            </div>
            <div className="rule-column">
              <h4>Épreuves Terminales (60%)</h4>
              <ul className="rule-list">
                <li className="rule-item"><span>Français Écrit (Épreuve de 1ère)</span><span className="rule-coef">Coef. 5</span></li>
                <li className="rule-item"><span>Français Oral (Épreuve de 1ère)</span><span className="rule-coef">Coef. 5</span></li>
                <li className="rule-item"><span>Épreuve écrite Spécialité 1</span><span className="rule-coef">Coef. 16</span></li>
                <li className="rule-item"><span>Épreuve écrite Spécialité 2</span><span className="rule-coef">Coef. 16</span></li>
                <li className="rule-item"><span>Philosophie (Épreuve écrite)</span><span className="rule-coef">Coef. 8</span></li>
                <li className="rule-item"><span>Grand Oral</span><span className="rule-coef">Coef. 10</span></li>
              </ul>
            </div>
          </div>
        )}
      </div>

      {/* FILTER BAR */}
      <section className="filter-bar">
        <div className="filter-group">
          <label className="filter-label">Recherche</label>
          <div className="input-search-container">
            <Search className="input-search-icon" size={16} />
            <input 
              type="text" 
              placeholder="Nom ou prénom..." 
              className="form-control form-control-search"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
            />
          </div>
        </div>

        <div className="filter-group">
          <label className="filter-label">Spécialité</label>
          <select 
            className="form-control"
            value={filterSpec}
            onChange={(e) => setFilterSpec(e.target.value)}
          >
            <option value="">Toutes les spécialités</option>
            {Object.keys(SPEC_SHORT_LABELS).map(sp => (
              <option key={sp} value={sp}>{SPEC_SHORT_LABELS[sp]}</option>
            ))}
          </select>
        </div>

        <div className="filter-group">
          <label className="filter-label">Statut</label>
          <select 
            className="form-control"
            value={filterStatus}
            onChange={(e) => setFilterStatus(e.target.value)}
          >
            <option value="">Tous les statuts</option>
            <option value="Admis">Admis</option>
            <option value="Rattrapage">Rattrapage</option>
            <option value="Refusé">Refusé</option>
          </select>
        </div>

        <div className="filter-group">
          <label className="filter-label">Mention</label>
          <select 
            className="form-control"
            value={filterMention}
            onChange={(e) => setFilterMention(e.target.value)}
          >
            <option value="">Toutes les mentions</option>
            <option value="Très bien avec félicitations">Très bien avec fél.</option>
            <option value="Très bien">Très bien</option>
            <option value="Bien">Bien</option>
            <option value="Assez bien">Assez bien</option>
            <option value="Admis sans mention">Sans mention</option>
          </select>
        </div>
      </section>

      {/* CANDIDATES TABLE */}
      <section className="table-card">
        <div className="table-header-info">
          <div>Candidats ({filteredStudents.length} sur {students.length} - cliquez une ligne pour simuler)</div>
          <div>Données réelles LFJP 2026</div>
        </div>
        
        <div className="table-container">
          <table className="cohort-table">
            <thead>
              <tr>
                <th className="sortable" onClick={() => toggleSort('name')}>Nom {sortField === 'name' ? (sortDirection === 'asc' ? '▲' : '▼') : ''}</th>
                <th>Prénom</th>
                <th>Spécialités Terminale</th>
                <th className="sortable" onClick={() => toggleSort('cc')}>Contrôle Continu (40%) {sortField === 'cc' ? (sortDirection === 'asc' ? '▲' : '▼') : ''}</th>
                <th className="sortable" onClick={() => toggleSort('terminal')}>Terminales (60%) {sortField === 'terminal' ? (sortDirection === 'asc' ? '▲' : '▼') : ''}</th>
                <th className="sortable" onClick={() => toggleSort('final')}>Moyenne finale {sortField === 'final' ? (sortDirection === 'asc' ? '▲' : '▼') : ''}</th>
                <th>Statut</th>
                <th>Mention</th>
              </tr>
            </thead>
            <tbody>
              {filteredStudents.map(student => {
                const status = student.results.status
                let statusClass = 'badge-admis'
                if (status === 'Rattrapage') statusClass = 'badge-rattrapage'
                if (status === 'Refusé') statusClass = 'badge-refuse'
                
                let isMentionHighlight = student.results.mention !== 'Admis sans mention' && student.results.mention !== 'Aucune'
                
                return (
                  <tr key={student.full_id} onClick={() => setSelectedStudentId(student.full_id)}>
                    <td className="student-name">{student.family_name}</td>
                    <td className="student-firstname">{student.first_name}</td>
                    <td>
                      <div style={{ display: 'flex', gap: '0.35rem', flexWrap: 'wrap', alignItems: 'center' }}>
                        {student.specialties.map(sp => (
                          <span key={sp} className="badge badge-spec" title={sp}>
                            {SPEC_SHORT_LABELS[sp] || sp}
                          </span>
                        ))}
                      </div>
                    </td>
                    <td>{student.results.ccAverage.toFixed(2).replace('.', ',')}</td>
                    <td>{student.results.terminalAverage.toFixed(2).replace('.', ',')}</td>
                    <td className="student-name">{student.results.finalAverage.toFixed(2).replace('.', ',')}</td>
                    <td>
                      <span className={`badge ${statusClass}`}>{status}</span>
                    </td>
                    <td className={`mention-text ${isMentionHighlight ? 'mention-highlight' : ''}`}>
                      {student.results.mention}
                    </td>
                  </tr>
                )
              })}
              
              {filteredStudents.length === 0 && (
                <tr>
                  <td colSpan="8" style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-muted)' }}>
                    Aucun candidat trouvé pour les critères de recherche actuels.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      {/* SIMULATION DIALOG / MODAL (SIDE DRAWER) */}
      {selectedStudentId && activeStudent && activeSim && (
        <div className="modal-overlay" onClick={() => setSelectedStudentId(null)}>
          <div className="modal-drawer" onClick={(e) => e.stopPropagation()}>
            {/* Modal Header */}
            <div className="modal-header">
              <div className="modal-header-meta">Fiche de simulation • Candidat</div>
              <h2 className="modal-title">{activeStudent.name}</h2>
              <div className="modal-subtitle">
                Classe Terminale Générale • N° {activeStudent.candidate_id} • Né(e) le {activeStudent.dob}
              </div>
              <button className="modal-close-btn" onClick={() => setSelectedStudentId(null)} title="Fermer">
                <X size={20} />
              </button>
            </div>

            {/* Modal Body */}
            <div className="modal-body">
              {/* Section 1: Contrôle Continu (40%) */}
              <div className="modal-section">
                <div className="modal-section-title-bar">
                  <h3 className="modal-section-title">Contrôle continu</h3>
                  <div className="modal-section-aside">Moyenne: <strong>{activeStudent.results.ccAverage.toFixed(2).replace('.', ',')} / 20</strong></div>
                </div>
                
                <div className="grades-columns-container">
                  {/* Column 1: Classe de 1ère */}
                  <div className="grades-column">
                    <div className="grades-column-header">
                      <span>Classe de 1ère</span>
                      <span>Coef. 21</span>
                    </div>
                    
                    <div className="grade-row">
                      <span className="grade-label">Histoire-Géographie</span>
                      <span className="grade-val">{activeStudent.grades['LP001']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">LVA (Anglais)</span>
                      <span className="grade-val">{activeStudent.grades['LP003']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">LVB (Espagnol)</span>
                      <span className="grade-val">{activeStudent.grades['LP004']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">Ens. Scientifique</span>
                      <span className="grade-val">
                        {(activeStudent.grades['LP005']?.value ?? activeStudent.grades['LP007']?.value)?.toFixed(2).replace('.', ',') || "10,00"}
                      </span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">EMC</span>
                      <span className="grade-val">{activeStudent.grades['LP002']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row" style={{ borderBottomColor: 'transparent' }}>
                      <span className="grade-label-highlight" title={activeStudent.dropped_specialty.label}>
                        Spé non poursuivie : {SPEC_SHORT_LABELS[activeStudent.dropped_specialty.label] || activeStudent.dropped_specialty.label || "Spé 1ère"}
                      </span>
                      <span className="grade-val">{activeStudent.dropped_specialty.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>

                  </div>
                  
                  {/* Column 2: Classe de Terminale */}
                  <div className="grades-column">
                    <div className="grades-column-header">
                      <span>Classe de Terminale</span>
                      <span>Coef. 19</span>
                    </div>
                    
                    <div className="grade-row">
                      <span className="grade-label">Histoire-Géographie</span>
                      <span className="grade-val">{activeStudent.grades['LT001']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">LVA (Anglais)</span>
                      <span className="grade-val">{activeStudent.grades['LT003']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">LVB (Espagnol)</span>
                      <span className="grade-val">{activeStudent.grades['LT004']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">Ens. Scientifique</span>
                      <span className="grade-val">{activeStudent.grades['LT005']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    <div className="grade-row">
                      <span className="grade-label">EMC</span>
                      <span className="grade-val">{activeStudent.grades['LT002']?.value?.toFixed(2).replace('.', ',') || "10,00"}</span>
                    </div>
                    
                    {/* EPS FIELD */}
                    {activeStudent.eps_grade === 'dispense' ? (
                      <div className="grade-row" style={{ borderBottomColor: 'var(--danger-text)' }}>
                        <span className="grade-label-highlight" style={{ color: 'var(--danger-text)' }}>EPS (Dispensé)</span>
                        <span className="badge-real" style={{ backgroundColor: 'var(--danger-text)', color: 'white' }}>Dispensé</span>
                      </div>
                    ) : activeStudent.eps_grade !== undefined ? (
                      <div className="grade-row" style={{ borderBottomColor: 'var(--success-text)' }}>
                        <span className="grade-label-highlight" style={{ color: 'var(--success-text)' }}>EPS (Note Réelle)</span>
                        <span className="grade-val" style={{ color: 'var(--success-text)' }}>{activeStudent.eps_grade.toFixed(2).replace('.', ',')}</span>
                      </div>
                    ) : (
                      <div className="grade-row" style={{ borderBottomColor: 'var(--primary)' }}>
                        <span className="grade-label-highlight" style={{ color: 'var(--primary)' }}>EPS (Simulé - Coef 6)</span>
                        <span className="grade-val" style={{ color: 'var(--primary)' }}>{activeSim.eps.toFixed(2).replace('.', ',')}</span>
                      </div>
                    )}
                  </div>
                </div>
                
                {/* EPS SLIDER CONTROL */}
                {activeStudent.eps_grade === undefined ? (
                  <div style={{ marginTop: '0.75rem', display: 'flex', gap: '1rem', alignItems: 'center' }}>
                    <span className="sim-coef" style={{ minWidth: '100px' }}>Simuler l'EPS:</span>
                    <input 
                      type="range" 
                      min="0" 
                      max="20" 
                      step="0.1" 
                      className="slider"
                      value={activeSim.eps}
                      onChange={(e) => handleGradeChange(activeStudent.full_id, 'eps', e.target.value)}
                    />
                    <input 
                      type="number" 
                      min="0" 
                      max="20" 
                      step="0.1"
                      className="sim-num-input"
                      value={activeSim.eps}
                      onChange={(e) => handleGradeChange(activeStudent.full_id, 'eps', e.target.value)}
                    />
                  </div>
                ) : activeStudent.eps_grade === 'dispense' ? (
                  <div style={{ marginTop: '0.75rem', padding: '0.5rem 0.75rem', fontSize: '0.8rem', backgroundColor: 'var(--bg-main)', borderLeft: '3px solid var(--danger-text)', borderRadius: '0.25rem' }}>
                    <strong>Note d'exemption :</strong> Cet élève est dispensé d'EPS. Le coefficient total du Contrôle Continu est de <strong>34</strong> et le coefficient global du Baccalauréat est de <strong>94</strong> au lieu de 100.
                  </div>
                ) : (
                  <div style={{ marginTop: '0.75rem', padding: '0.5rem 0.75rem', fontSize: '0.8rem', backgroundColor: 'var(--bg-main)', borderLeft: '3px solid var(--success-text)', borderRadius: '0.25rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span><strong>Note réelle verrouillée :</strong> Évaluation d'EPS sur l'année de {activeStudent.eps_grade.toFixed(2).replace('.', ',')}/20.</span>
                    <span className="badge-real" style={{ alignSelf: 'center', marginLeft: '0.5rem' }}>Note réelle</span>
                  </div>
                )}
                <div className="modal-section-description">Moyenne pondérée du contrôle continu, soit 40 % de la note finale.</div>
              </div>

              {/* Section 2: Épreuves Anticipées (Français) */}
              <div className="modal-section">
                <div className="modal-section-title-bar">
                  <h3 className="modal-section-title">Épreuves anticipées (Français)</h3>
                  <div className="modal-section-aside">Coef. 5 chacun</div>
                </div>

                <div className="grades-grid" style={{ marginBottom: activeSim.isFrenchMissing ? '0.75rem' : '0' }}>
                  {/* FRENCH WRITTEN */}
                  <div>
                    <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '0.25rem' }}>Français Écrit (coef 5)</div>
                    {!activeSim.isFrenchMissing ? (
                      <div className="locked-box">
                        <span className="grade-val">{activeStudent.french_written?.toFixed(2).replace('.', ',')}</span>
                        <span className="badge-real">Note réelle</span>
                      </div>
                    ) : (
                      <div className="sim-input-wrapper">
                        <input 
                          type="range" 
                          min="0" 
                          max="20" 
                          step="0.1" 
                          className="slider"
                          value={activeSim.frenchWritten}
                          onChange={(e) => handleGradeChange(activeStudent.full_id, 'frenchWritten', e.target.value)}
                        />
                        <input 
                          type="number" 
                          min="0" 
                          max="20" 
                          step="0.1"
                          className="sim-num-input"
                          value={activeSim.frenchWritten}
                          onChange={(e) => handleGradeChange(activeStudent.full_id, 'frenchWritten', e.target.value)}
                        />
                      </div>
                    )}
                  </div>

                  {/* FRENCH ORAL */}
                  <div>
                    <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '0.25rem' }}>Français Oral (coef 5)</div>
                    {!activeSim.isFrenchMissing ? (
                      <div className="locked-box">
                        <span className="grade-val">{activeStudent.french_oral?.toFixed(2).replace('.', ',')}</span>
                        <span className="badge-real">Note réelle</span>
                      </div>
                    ) : (
                      <div className="sim-input-wrapper">
                        <input 
                          type="range" 
                          min="0" 
                          max="20" 
                          step="0.1" 
                          className="slider"
                          value={activeSim.frenchOral}
                          onChange={(e) => handleGradeChange(activeStudent.full_id, 'frenchOral', e.target.value)}
                        />
                        <input 
                          type="number" 
                          min="0" 
                          max="20" 
                          step="0.1"
                          className="sim-num-input"
                          value={activeSim.frenchOral}
                          onChange={(e) => handleGradeChange(activeStudent.full_id, 'frenchOral', e.target.value)}
                        />
                      </div>
                    )}
                  </div>
                </div>
                {activeSim.isFrenchMissing && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginTop: '0.5rem' }}>
                    <span className="badge-warning-custom">Non saisie</span>
                    <span style={{ fontSize: '0.78rem', color: 'var(--danger-text)', fontWeight: 600 }}>
                      Note de français non présente dans le dossier de 1ère. Saisie manuelle requise !
                    </span>
                  </div>
                )}
              </div>

              {/* Section 3: Épreuves Terminales (60% - remaining exams) */}
              <div className="modal-section">
                <div className="modal-section-title-bar">
                  <h3 className="modal-section-title">Épreuves terminales</h3>
                  <div className="modal-section-aside">Notes simulées, calibrées à 10/20</div>
                </div>

                <div className="sim-grid">
                  {/* SPECIALTY 1 */}
                  <div className="sim-group">
                    <div className="sim-label-area">
                      <span className="sim-label" title={activeStudent.specialties[0]}>
                        Spé 1 : {SPEC_SHORT_LABELS[activeStudent.specialties[0]] || activeStudent.specialties[0] || "Spécialité 1"}
                      </span>
                      <span className="sim-coef">Coef. 16</span>
                    </div>
                    <div className="sim-input-wrapper">
                      <input 
                        type="range" 
                        min="0" 
                        max="20" 
                        step="0.1" 
                        className="slider"
                        value={activeSim.spe1}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'spe1', e.target.value)}
                      />
                      <input 
                        type="number" 
                        min="0" 
                        max="20" 
                        step="0.1"
                        className="sim-num-input"
                        value={activeSim.spe1}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'spe1', e.target.value)}
                      />
                    </div>
                  </div>

                  {/* SPECIALTY 2 */}
                  <div className="sim-group">
                    <div className="sim-label-area">
                      <span className="sim-label" title={activeStudent.specialties[1]}>
                        Spé 2 : {SPEC_SHORT_LABELS[activeStudent.specialties[1]] || activeStudent.specialties[1] || "Spécialité 2"}
                      </span>
                      <span className="sim-coef">Coef. 16</span>
                    </div>
                    <div className="sim-input-wrapper">
                      <input 
                        type="range" 
                        min="0" 
                        max="20" 
                        step="0.1" 
                        className="slider"
                        value={activeSim.spe2}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'spe2', e.target.value)}
                      />
                      <input 
                        type="number" 
                        min="0" 
                        max="20" 
                        step="0.1"
                        className="sim-num-input"
                        value={activeSim.spe2}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'spe2', e.target.value)}
                      />
                    </div>
                  </div>

                  {/* PHILOSOPHY */}
                  <div className="sim-group">
                    <div className="sim-label-area">
                      <span className="sim-label">Philosophie (Écrit)</span>
                      <span className="sim-coef">Coef. 8</span>
                    </div>
                    <div className="sim-input-wrapper">
                      <input 
                        type="range" 
                        min="0" 
                        max="20" 
                        step="0.1" 
                        className="slider"
                        value={activeSim.philo}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'philo', e.target.value)}
                      />
                      <input 
                        type="number" 
                        min="0" 
                        max="20" 
                        step="0.1"
                        className="sim-num-input"
                        value={activeSim.philo}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'philo', e.target.value)}
                      />
                    </div>
                  </div>

                  {/* GRAND ORAL */}
                  <div className="sim-group">
                    <div className="sim-label-area">
                      <span className="sim-label">Grand Oral</span>
                      <span className="sim-coef">Coef. 10</span>
                    </div>
                    <div className="sim-input-wrapper">
                      <input 
                        type="range" 
                        min="0" 
                        max="20" 
                        step="0.1" 
                        className="slider"
                        value={activeSim.grandOral}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'grandOral', e.target.value)}
                      />
                      <input 
                        type="number" 
                        min="0" 
                        max="20" 
                        step="0.1"
                        className="sim-num-input"
                        value={activeSim.grandOral}
                        onChange={(e) => handleGradeChange(activeStudent.full_id, 'grandOral', e.target.value)}
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* LIVE SIMULATED AVERAGE RESULT BOX */}
              <div className="calc-summary-bar">
                <div className="calc-summary-item">
                  <span className="calc-summary-label">Moy. Terminales (60%)</span>
                  <span className="calc-summary-val">{activeStudent.results.terminalAverage.toFixed(2).replace('.', ',')}</span>
                </div>
                
                <div className="calc-summary-item" style={{ borderLeft: '1px solid var(--border-medium)', paddingLeft: '1.5rem' }}>
                  <span className="calc-summary-label">Moyenne Générale</span>
                  <span className="calc-summary-val-giant">{activeStudent.results.finalAverage.toFixed(2).replace('.', ',')}</span>
                </div>

                <div className="calc-summary-item" style={{ alignItems: 'flex-end', marginLeft: 'auto' }}>
                  <span className="calc-summary-label" style={{ marginBottom: '0.25rem' }}>Résultat</span>
                  <span className={`badge ${
                    activeStudent.results.status === 'Admis' ? 'badge-admis' : 
                    activeStudent.results.status === 'Rattrapage' ? 'badge-rattrapage' : 'badge-refuse'
                  }`} style={{ fontSize: '0.95rem', padding: '0.35rem 0.85rem' }}>
                    {activeStudent.results.status}
                  </span>
                  <span style={{ fontSize: '0.78rem', color: 'var(--text-muted)', fontWeight: 600, marginTop: '0.25rem' }}>
                    {activeStudent.results.mention}
                  </span>
                </div>
              </div>

              {/* Section 4: Moyenne Terminale Minimale */}
              <div className="modal-section">
                <div className="modal-section-title-bar">
                  <h3 className="modal-section-title">Moyenne épreuves terminales minimale</h3>
                  <div className="modal-section-aside">Seuils requis</div>
                </div>
                
                <table className="min-table">
                  <tbody>
                    {minRequiredAverages.map(item => {
                      let valClass = ''
                      if (item.status === 'acquired') valClass = 'acquired'
                      if (item.status === 'impossible') valClass = 'impossible'
                      
                      return (
                        <tr key={item.target}>
                          <td className="min-label">{item.label}</td>
                          <td className="min-target">Note générale ≥ {item.target.toFixed(2).replace('.', ',')}</td>
                          <td className={`min-val ${valClass}`}>{item.note}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
                <div className="modal-section-description">
                  Moyenne arithmétique pondérée requise sur les 4 épreuves terminales restantes (Spé 1, Spé 2, Philosophie, Grand Oral - coeff 50) pour atteindre le palier correspondant.
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="modal-footer">
              <button className="btn" onClick={() => resetStudentSimulation(activeStudent.full_id)}>
                <RotateCcw size={16} />
                Réinitialiser à 10/20
              </button>
              <button 
                className="btn btn-primary" 
                onClick={() => exportIndividualToExcel(activeStudent, activeSim, activeStudent.results)}
              >
                <FileSpreadsheet size={16} />
                Fiche Excel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default App
