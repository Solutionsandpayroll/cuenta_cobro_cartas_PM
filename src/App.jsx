import { useState, useRef, useEffect } from 'react'
import { createPortal } from 'react-dom'
import * as XLSX from 'xlsx'
import PizZip from 'pizzip'
import { renderAsync } from 'docx-preview'
import html2canvas from 'html2canvas'
import { jsPDF } from 'jspdf'
import './App.css'

// ── JSONBin.io — persistencia del consecutivo en la nube ──────────────────
// Los valores se configuran en el archivo .env (ver .env.example)
const _JB_BIN_ID = import.meta.env.VITE_JSONBIN_BIN_ID || ''
const _JB_URL = `https://api.jsonbin.io/v3/b/${_JB_BIN_ID}`
const _JB_KEY = _JB_BIN_ID ? atob(import.meta.env.VITE_JSONBIN_API_KEY_B64 || '') : ''

// ── Conversión numérico → letras (español, formato colombiano) ──
const _unidades = ['','UN','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE',
  'DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS','DIECISIETE','DIECIOCHO','DIECINUEVE','VEINTE']
const _veintis   = ['','VEINTIÚN','VEINTIDÓS','VEINTITRÉS','VEINTICUATRO','VEINTICINCO',
  'VEINTISÉIS','VEINTISIETE','VEINTIOCHO','VEINTINUEVE']
const _decenas   = ['','','','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA']
const _centenas  = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
  'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS']

function _dec(n) {
  if (n <= 20) return _unidades[n]
  if (n < 30)  return _veintis[n - 20]
  const u = n % 10
  return u === 0 ? _decenas[Math.floor(n/10)] : _decenas[Math.floor(n/10)] + ' Y ' + _unidades[u]
}
function _cen(n) {
  if (n === 0)   return ''
  if (n === 100) return 'CIEN'
  const resto = n % 100
  return _centenas[Math.floor(n/100)] + (resto ? ' ' + _dec(resto) : '')
}
function _conv(n) {
  if (n === 0)           return ''
  if (n < 100)           return _dec(n)
  if (n < 1000)          return _cen(n)
  if (n < 1000000) {
    const m = Math.floor(n/1000), r = n % 1000
    const pre = m === 1 ? 'MIL' : _conv(m) + ' MIL'
    return pre + (r ? ' ' + _cen(r) : '')
  }
  if (n < 1000000000) {
    const m = Math.floor(n/1000000), r = n % 1000000
    const pre = m === 1 ? 'UN MILLÓN' : _conv(m) + ' MILLONES'
    return pre + (r ? ' ' + _conv(r) : '')
  }
  const m = Math.floor(n/1000000000), r = n % 1000000000
  const pre = m === 1 ? 'MIL MILLONES' : _conv(m) + ' MIL MILLONES'
  return pre + (r ? ' ' + _conv(r) : '')
}
function numeroALetras(valorStr) {
  // Acepta formatos colombianos: $1.500.000 / 1500000 / 1,500,000
  const cleaned = String(valorStr).replace(/[$\s]/g, '').replace(/\./g, '').replace(/,/g, '.')
  const num = parseInt(cleaned, 10)
  if (isNaN(num) || cleaned === '') return valorStr
  if (num === 0) return 'CERO PESOS M/CTE'
  return _conv(num).trim() + ' PESOS M/CTE'
}
function parseValorCOP(valorStr) {
  // Parsea formato colombiano: "$1.500.000" / "1500000" → número
  const cleaned = String(valorStr).replace(/[$\s]/g, '').replace(/\./g, '').replace(/,/g, '.')
  const num = parseFloat(cleaned)
  return isNaN(num) ? 0 : num
}
function formatValorCOP(num) {
  // Formatea número como "$ 54.564" (puntos como separadores de miles)
  const rounded = Math.round(num)
  return '$ ' + rounded.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.')
}

function App() {
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [activeTab, setActiveTab] = useState('archivos') // 'archivos' | 'editor'
  const [showBlockedModal, setShowBlockedModal] = useState(false)

  // Base Empleados
  const [baseEmpleadosFile, setBaseEmpleadosFile] = useState(null)
  const [dragEmpleados, setDragEmpleados] = useState(false)
  const inputEmpleadosRef = useRef(null)

  // Base ReteICA
  const [baseReteICAFile, setBaseReteICAFile] = useState(null)
  const [dragReteICA, setDragReteICA] = useState(false)
  const inputReteICARef = useRef(null)
  const [reteICAData, setReteICAData] = useState([])
  const [selectedCodigo, setSelectedCodigo] = useState('')
  const [selectedRow, setSelectedRow] = useState(null)
  const [reteICASource, setReteICASource] = useState('integrada') // 'integrada' | 'usuario'

  // Base Empleados - datos parseados
  const [empleadosData, setEmpleadosData] = useState([])
  const [empleadoBusqueda, setEmpleadoBusqueda] = useState('')
  const [empleadoDropdownVisible, setEmpleadoDropdownVisible] = useState(false)
  const [selectedEmpleado, setSelectedEmpleado] = useState(null)
  const empleadoInputRef = useRef(null)
  const [dropdownPortalPos, setDropdownPortalPos] = useState({ top: 0, left: 0, width: 0 })

  const recalcDropdownPos = () => {
    if (empleadoInputRef.current) {
      const r = empleadoInputRef.current.getBoundingClientRect()
      setDropdownPortalPos({ top: r.bottom + 4, left: r.left, width: r.width })
    }
  }
  // Campos del formulario posterior a selección
  const [tipoDoc, setTipoDoc] = useState('')
  const [fechaProceso, setFechaProceso] = useState('')
  const [tieneReteICA, setTieneReteICA] = useState('')
  const [tieneRetencionFuente, setTieneRetencionFuente] = useState('')
  const [valorRetencionFuente, setValorRetencionFuente] = useState('')
  const [concepto, setConcepto] = useState('')
  const [valor, setValor] = useState('')
  const [consecutivo, setConsecutivo] = useState(() => {
    const saved = localStorage.getItem('pm_cartas_consecutivo')
    return saved !== null ? parseInt(saved, 10) : 16
  })
  const [generando, setGenerando] = useState(false)
  const [generandoPDF, setGenerandoPDF] = useState(false)

  // Cargar el consecutivo desde JSONBin al iniciar
  useEffect(() => {
    if (_JB_BIN_ID && _JB_KEY) {
      fetch(`${_JB_URL}/latest`, { headers: { 'X-Master-Key': _JB_KEY } })
        .then(r => r.json())
        .then(data => {
          if (data?.record?.consecutivo != null) {
            setConsecutivo(data.record.consecutivo)
          }
        })
        .catch(e => console.warn('[JSONBin] No se pudo leer el consecutivo:', e))
    }
  }, [])

  // Guarda el consecutivo en JSONBin y también en localStorage como respaldo
  const saveConsecutivo = async (n) => {
    localStorage.setItem('pm_cartas_consecutivo', String(n))
    if (_JB_BIN_ID && _JB_KEY) {
      try {
        await fetch(_JB_URL, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json', 'X-Master-Key': _JB_KEY },
          body: JSON.stringify({ consecutivo: n }),
        })
      } catch (e) {
        console.warn('[JSONBin] No se pudo guardar el consecutivo:', e)
      }
    }
  }

  const formatFecha = (isoDate) => {
    if (!isoDate) return ''
    const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
    const [year, month, day] = isoDate.split('-').map(Number)
    return `${day} de ${meses[month - 1]} de ${year}`
  }

  const buildFilledDocxBuf = async () => {
    const res = await fetch('/Carta cuenta de cobro.docx')
    const buf = await res.arrayBuffer()
    const zip = new PizZip(buf)
    const xmlFile = zip.file('word/document.xml')
    if (!xmlFile) throw new Error('No se encontró word/document.xml')
    let xml = xmlFile.asText()
    const replace = (placeholder, value) => {
      xml = xml.split(placeholder).join(value ?? '')
    }
    replace('@@FECHA@@', formatFecha(fechaProceso))
    replace('@@NOMBRECOMPLETO@@', selectedEmpleado?.nombreCompleto ?? '')
    replace('@@NUMERODOCUMENTO@@', selectedEmpleado?.nit ?? '')
    replace('@@TELEFONO@@', selectedEmpleado?.telefono ?? '')
    replace('@@DIRECCION@@', selectedEmpleado?.direccion ?? '')
    replace('@@DOCUMENTO@@', tipoDoc)
    replace('@@VALORRETEF@@', tieneRetencionFuente === 'no' ? 'No Aplica' : valorRetencionFuente)
    replace('@@VALORRETEICA@@', tieneReteICA === 'no' ? 'No Aplica' : (selectedRow?.tarifaPorMil != null ? selectedRow.tarifaPorMil + '/1000' : ''))
    replace('@@CODIGOACTIVIDADECONOMICA@@', selectedCodigo)
    const _numValor = parseValorCOP(valor)
    const _valorICA = tieneReteICA === 'no' ? 0 : (selectedRow?.tarifaPorMil ?? 0) / 1000 * _numValor
    replace('@@VALORICA@@', tieneReteICA === 'no' ? '-' : formatValorCOP(_valorICA))
    const _pctFuente = parseFloat(String(valorRetencionFuente).replace('%', '').trim())
    const _valorFuente = tieneRetencionFuente === 'no' ? 0 : (isNaN(_pctFuente) ? 0 : _pctFuente / 100) * _numValor
    replace('@@VALORFUENTE@@', tieneRetencionFuente === 'no' ? '-' : formatValorCOP(_valorFuente))
    replace('@@TOTALPAGAR@@', formatValorCOP(_numValor - _valorICA - _valorFuente))
    if (tieneReteICA === 'no') {
      xml = xml.replace(/@@INICIOICA@@[\s\S]*?@@FINICA@@/g, '')
    } else {
      replace('@@INICIOICA@@', '')
      replace('@@FINICA@@', '')
    }
    replace('@@CONCEPTO@@', concepto)
    replace('@@VALOR@@', valor)
    replace('@@VALORENLETRAS@@', numeroALetras(valor))
    replace('@@CONSECUTIVO@@', String(consecutivo))
    zip.file('word/document.xml', xml)
    const arrayBuffer = zip.generate({ type: 'arraybuffer' })
    const nombre = selectedEmpleado?.nombreCompleto
      ? `Carta - ${selectedEmpleado.nombreCompleto}`
      : 'Carta cuenta de cobro'
    return { arrayBuffer, nombre }
  }

  const generateCarta = async () => {
    setGenerando(true)
    try {
      const { arrayBuffer, nombre } = await buildFilledDocxBuf()
      const blob = new Blob([arrayBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = nombre + '.docx'
      a.click()
      URL.revokeObjectURL(url)
      const next = consecutivo + 1
      setConsecutivo(next)
      await saveConsecutivo(next)
    } catch (err) {
      console.error(err)
      alert('Error al generar la carta: ' + err.message)
    } finally {
      setGenerando(false)
    }
  }

  const generatePDF = async () => {
    setGenerandoPDF(true)
    try {
      const { arrayBuffer, nombre } = await buildFilledDocxBuf()

      // 794px = A4 a 96dpi. Sin este ancho, docx-preview renderiza en un ancho
      // arbitrario y al escalar al PDF el contenido aparece "con zoom".
      const container = document.createElement('div')
      container.style.cssText = 'position:absolute;top:-9999px;left:0;width:794px;background:white;'
      document.body.appendChild(container)

      // Alias de fuentes: "Arial MT" no es reconocida por el navegador como Arial,
      // causando fallback a otra fuente. Registramos @font-face antes de renderizar.
      const fontAliases = document.createElement('style')
      fontAliases.textContent = [
        "@font-face { font-family: 'Arial MT';  src: local('Arial'), local('ArialMT'); }",
        "@font-face { font-family: 'ArialMT';   src: local('Arial'), local('Arial MT'); }",
        "@font-face { font-family: 'Arial MT Bold'; src: local('Arial Bold'), local('Arial-BoldMT'); }",
      ].join('\n')
      document.head.appendChild(fontAliases)

      // Renderizar el .docx con fidelidad visual completa
      await renderAsync(arrayBuffer, container, null, {
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
        experimental: true,
      })

      // Buscar las páginas renderizadas (docx-preview las crea como <section>)
      const pages = Array.from(container.querySelectorAll('section'))
      const targets = pages.length > 0 ? pages : [container]

      let pdf = null

      for (let i = 0; i < targets.length; i++) {
        const el = targets[i]
        const elW = el.offsetWidth  || 794
        const elH = el.offsetHeight || 1122

        const canvas = await html2canvas(el, {
          scale: 2,
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff',
          allowTaint: true,
          width: elW,
          height: elH,
          x: 0,
          y: 0,
          scrollX: 0,
          scrollY: 0,
        })

        // Dimensiones del PDF derivadas de los píxeles reales (px→mm a 96dpi)
        // así no hay distorsión de escala independientemente del tamaño de página del .docx
        const pageMmW = (elW / 96) * 25.4
        const pageMmH = (elH / 96) * 25.4
        const imgData  = canvas.toDataURL('image/jpeg', 0.97)

        if (i === 0) {
          pdf = new jsPDF({
            orientation: pageMmW > pageMmH ? 'landscape' : 'portrait',
            unit: 'mm',
            format: [pageMmW, pageMmH],
          })
        } else {
          pdf.addPage([pageMmW, pageMmH])
        }
        pdf.addImage(imgData, 'JPEG', 0, 0, pageMmW, pageMmH)
      }

      document.body.removeChild(container)
      document.head.removeChild(fontAliases)
      pdf.save(nombre + '.pdf')

      const next = consecutivo + 1
      setConsecutivo(next)
      await saveConsecutivo(next)
    } catch (err) {
      console.error(err)
      alert('Error al generar el PDF: ' + err.message)
    } finally {
      setGenerandoPDF(false)
    }
  }

  const formatFileSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B'
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB'
  }

  const parseSheetData = (arrayBuffer) => {
    const data = new Uint8Array(arrayBuffer)
    const workbook = XLSX.read(data, { type: 'array' })
    const sheet = workbook.Sheets['GRAVADAS']
    if (!sheet) {
      alert('No se encontró la hoja "GRAVADAS" en el archivo.')
      return []
    }
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 })
    const parsed = []
    for (let i = 2; i < rows.length; i++) {
      const row = rows[i]
      const codigo = row[1]
      const descripcion = row[2]
      const tarifaPorMil = row[6]
      if (codigo !== undefined && codigo !== null && String(codigo).trim() !== '') {
        parsed.push({
          codigo: String(codigo).trim(),
          descripcion: descripcion !== undefined ? String(descripcion) : '',
          tarifaPorMil: tarifaPorMil !== undefined ? tarifaPorMil : '',
        })
      }
    }
    return parsed
  }

  // Cargar el Excel integrado al montar el componente
  useEffect(() => {
    fetch('/Base ReteICA.xlsx')
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const parsed = parseSheetData(buffer)
        setReteICAData(parsed)
      })
      .catch(() => {
        console.warn('No se pudo cargar la Base ReteICA integrada.')
      })
  }, [])

  const getTodayISO = () => new Date().toISOString().split('T')[0]

  const resetEmpleadoForm = () => {
    setSelectedEmpleado(null)
    setEmpleadoBusqueda('')
    setEmpleadoDropdownVisible(false)
    setTipoDoc('')
    setFechaProceso('')
    setTieneReteICA('')
    setTieneRetencionFuente('')
    setValorRetencionFuente('')
    setConcepto('')
  }

  const parseEmpleadosData = (arrayBuffer) => {
    const data = new Uint8Array(arrayBuffer)
    const workbook = XLSX.read(data, { type: 'array' })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 })
    // Fila 3 (índice 2) = encabezados, datos desde fila 4 (índice 3)
    const parsed = []
    for (let i = 3; i < rows.length; i++) {
      const row = rows[i]
      const nit          = row[0]  // A: Nit Identificacion
      const nombreCompleto = row[1] // B: Nombre Completo
      const tipoDocVal   = row[4]  // E: Tipo Doc
      const telefono     = row[5]  // F: Telefono
      const direccion    = row[6]  // G: Direccion
      const departamento = row[7]  // H: Departamento
      if (nit !== undefined && nit !== null && String(nit).trim() !== '') {
        parsed.push({
          nit: String(nit).trim(),
          nombreCompleto: nombreCompleto ? String(nombreCompleto).trim() : '',
          tipoDoc: tipoDocVal ? String(tipoDocVal).trim() : '',
          telefono: telefono ? String(telefono).trim() : '',
          direccion: direccion ? String(direccion).trim() : '',
          departamento: departamento ? String(departamento).trim() : '',
        })
      }
    }
    return parsed
  }

  const handleEmpleadosFile = (file) => {
    if (!file) return
    setBaseEmpleadosFile(file)
    setEmpleadosData([])
    resetEmpleadoForm()
    const reader = new FileReader()
    reader.onload = (e) => {
      const parsed = parseEmpleadosData(e.target.result)
      setEmpleadosData(parsed)
    }
    reader.readAsArrayBuffer(file)
  }

  const handleRemoveEmpleados = () => {
    setBaseEmpleadosFile(null)
    if (inputEmpleadosRef.current) inputEmpleadosRef.current.value = ''
    setEmpleadosData([])
    resetEmpleadoForm()
  }

  const empleadosFiltrados = empleadoBusqueda.length > 0
    ? empleadosData.filter((e) => e.nit.startsWith(empleadoBusqueda)).slice(0, 25)
    : empleadosData.slice(0, 25)

  const handleSeleccionarEmpleado = (emp) => {
    setSelectedEmpleado(emp)
    setEmpleadoBusqueda(emp.nit)
    setEmpleadoDropdownVisible(false)
    setTipoDoc(emp.tipoDoc || '')
    setFechaProceso(getTodayISO())
    setTieneReteICA('')
    setTieneRetencionFuente('')
    setValorRetencionFuente('')
  }

  const handleReteICAFile = (file) => {
    if (!file) return
    setBaseReteICAFile(file)
    setSelectedCodigo('')
    setSelectedRow(null)
    setReteICAData([])
    const reader = new FileReader()
    reader.onload = (e) => {
      const parsed = parseSheetData(e.target.result)
      setReteICAData(parsed)
      setReteICASource('usuario')
    }
    reader.readAsArrayBuffer(file)
  }

  const handleRemoveReteICA = () => {
    setBaseReteICAFile(null)
    if (inputReteICARef.current) inputReteICARef.current.value = ''
    setSelectedCodigo('')
    setSelectedRow(null)
    setReteICASource('integrada')
    // Volver a cargar la base integrada
    fetch('/Base ReteICA.xlsx')
      .then((res) => res.arrayBuffer())
      .then((buffer) => {
        const parsed = parseSheetData(buffer)
        setReteICAData(parsed)
      })
  }

  const handleCodigoChange = (e) => {
    const val = e.target.value
    setSelectedCodigo(val)
    if (val === '') {
      setSelectedRow(null)
    } else {
      const found = reteICAData.find((r) => r.codigo === val)
      setSelectedRow(found || null)
    }
  }

  const ExcelIcon = () => (
    <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/>
      <line x1="8" y1="13" x2="16" y2="13"/>
      <line x1="8" y1="17" x2="16" y2="17"/>
      <line x1="10" y1="9" x2="8" y2="9"/>
    </svg>
  )

  const renderDropZone = (label, file, dragActive, inputRef, onFile, onDrag, onRemove, accept) => (
    <div className="form-group">
      <label className="label">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
        </svg>
        {label}
      </label>
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        className="file-input"
        onChange={(e) => onFile(e.target.files[0])}
      />
      {!file ? (
        <div
          className={`drop-zone ${dragActive ? 'drag-active' : ''}`}
          onClick={() => inputRef.current.click()}
          onDragOver={(e) => { e.preventDefault(); onDrag(true) }}
          onDragLeave={() => onDrag(false)}
          onDrop={(e) => { e.preventDefault(); onDrag(false); onFile(e.dataTransfer.files[0]) }}
        >
          <div className="drop-zone-content">
            <ExcelIcon />
            <div className="drop-zone-text">
              <span className="drop-zone-title">Arrastra el archivo aquí</span>
              <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
            </div>
            <span className="drop-zone-hint">Archivos .xlsx, .xls</span>
          </div>
        </div>
      ) : (
        <div className="drop-zone has-file">
          <div className="file-preview">
            <div className="file-icon">
              <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
              </svg>
            </div>
            <div className="file-details">
              <div className="file-name">{file.name}</div>
              <div className="file-size">{formatFileSize(file.size)}</div>
            </div>
            <button className="btn-remove" title="Quitar archivo" onClick={onRemove}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="18" y1="6" x2="6" y2="18"/>
                <line x1="6" y1="6" x2="18" y2="18"/>
              </svg>
            </button>
          </div>
        </div>
      )}
    </div>
  )

  return (
    <div className="app">
      {/* Header Corporativo Punto Medical */}
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img
                  src="/pm_logo.png"
                  alt="Punto Medical Logo"
                  width="60"
                  height="60"
                />
              </div>
              <div className="header-text">
                <h1>Punto Medical</h1>
                <p className="subtitle">Generación de Cartas - Cuenta de Cobro</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario</span>
            </div>
          </div>
        </div>
      </header>

      {/* Contenido Principal */}
      <main className="main-content">
        <div className="container">

          {/* Sección de ayuda */}
          <div className="help-section">
            <button
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" height="20" viewBox="0 0 24 24"
                fill="none" stroke="currentColor" strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Sube la Base de Empleados</strong>
                    <p>Carga el archivo Excel con la información de los empleados.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Configura el Editor de Información</strong>
                    <p>Selecciona el empleado, el código de actividad y completa los campos del proceso.</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Base ReteICA (opcional)</strong>
                    <p>Solo si los valores de tarifas han cambiado, sube una base actualizada.</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          {/* Tabs de navegación */}
          <div className="tabs-nav">
            <button
              className={`tab-btn ${activeTab === 'archivos' ? 'tab-btn--active' : ''}`}
              onClick={() => setActiveTab('archivos')}
            >
              <span className="tab-number">1</span>
              Carga de Archivos
            </button>
            <button
              className={`tab-btn ${activeTab === 'editor' ? 'tab-btn--active' : ''} ${empleadosData.length === 0 ? 'tab-btn--locked' : ''}`}
              onClick={() => {
                if (empleadosData.length === 0) {
                  setShowBlockedModal(true)
                } else {
                  setActiveTab('editor')
                }
              }}
            >
              <span className="tab-number">2</span>
              Editor de Información
              {empleadosData.length === 0 && (
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{marginLeft: '4px', opacity: 0.6}}>
                  <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
                  <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
                </svg>
              )}
            </button>
          </div>

          {/* Contenido de tabs */}
          <div className="tab-content">

            {/* ── Tab 1: Carga de Archivos ── */}
            {activeTab === 'archivos' && <section className="section-panel">
              <div className="section-header">
                <div className="section-number">1</div>
                <div>
                  <h2>Carga de Archivos</h2>
                  <p>Sube los archivos Excel necesarios</p>
                </div>
              </div>
              <div className="section-body">
                <div className="form-section">

                  {/* Base Empleados */}
                  {renderDropZone(
                    'Base empleados',
                    baseEmpleadosFile,
                    dragEmpleados,
                    inputEmpleadosRef,
                    handleEmpleadosFile,
                    setDragEmpleados,
                    handleRemoveEmpleados,
                    '.xlsx,.xls'
                  )}

                  {/* Base ReteICA (opcional) */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                      </svg>
                      Base ReteICA
                      <span className="badge-opcional">Opcional</span>
                    </label>
                    <p className="hint">Solo necesario si la base integrada está desactualizada (ej. cambios de tarifas).</p>
                    <input
                      ref={inputReteICARef}
                      type="file"
                      accept=".xlsx,.xls"
                      className="file-input"
                      onChange={(e) => handleReteICAFile(e.target.files[0])}
                    />
                    {!baseReteICAFile ? (
                      <div
                        className={`drop-zone ${dragReteICA ? 'drag-active' : ''}`}
                        onClick={() => inputReteICARef.current.click()}
                        onDragOver={(e) => { e.preventDefault(); setDragReteICA(true) }}
                        onDragLeave={() => setDragReteICA(false)}
                        onDrop={(e) => { e.preventDefault(); setDragReteICA(false); handleReteICAFile(e.dataTransfer.files[0]) }}
                      >
                        <div className="drop-zone-content">
                          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                            <line x1="8" y1="13" x2="16" y2="13"/>
                            <line x1="8" y1="17" x2="16" y2="17"/>
                            <line x1="10" y1="9" x2="8" y2="9"/>
                          </svg>
                          <div className="drop-zone-text">
                            <span className="drop-zone-title">Arrastra el archivo aquí</span>
                            <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
                          </div>
                          <span className="drop-zone-hint">Archivos .xlsx, .xls</span>
                        </div>
                      </div>
                    ) : (
                      <div className="drop-zone has-file">
                        <div className="file-preview">
                          <div className="file-icon">
                            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                              <polyline points="14 2 14 8 20 8"/>
                            </svg>
                          </div>
                          <div className="file-details">
                            <div className="file-name">{baseReteICAFile.name}</div>
                            <div className="file-size">{formatFileSize(baseReteICAFile.size)}</div>
                          </div>
                          <button className="btn-remove" title="Quitar archivo" onClick={handleRemoveReteICA}>
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <line x1="18" y1="6" x2="6" y2="18"/>
                              <line x1="6" y1="6" x2="18" y2="18"/>
                            </svg>
                          </button>
                        </div>
                      </div>
                    )}
                  </div>

                </div>
              </div>
            </section>}

            {/* ── Tab 2: Editor de Información ── */}
            {activeTab === 'editor' && <section className="section-panel">
              <div className="section-header">
                <div className="section-number">2</div>
                <div>
                  <h2>Editor de Información</h2>
                  <p>Configura los parámetros de la carta</p>
                </div>
              </div>
              <div className="section-body">
                <div className="form-section">

                  {/* 1. Fecha del Proceso */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <rect x="3" y="4" width="18" height="18" rx="2"/>
                        <line x1="16" y1="2" x2="16" y2="6"/>
                        <line x1="8" y1="2" x2="8" y2="6"/>
                        <line x1="3" y1="10" x2="21" y2="10"/>
                      </svg>
                      Fecha del Proceso
                    </label>
                    <input
                      type="date"
                      className="select-input"
                      value={fechaProceso}
                      onChange={(e) => setFechaProceso(e.target.value)}
                    />
                  </div>

                  {/* 2. Consecutivo */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <polyline points="9 11 12 14 22 4"/>
                        <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
                      </svg>
                      Consecutivo
                    </label>
                    <p className="hint">Se incrementa automáticamente al generar cada carta.</p>
                    <input
                      type="number"
                      className="select-input"
                      value={consecutivo}
                      min={1}
                      onChange={(e) => {
                        const v = parseInt(e.target.value, 10)
                        if (!isNaN(v)) {
                          setConsecutivo(v)
                          localStorage.setItem('pm_cartas_consecutivo', String(v))
                        }
                      }}
                    />
                  </div>

                  {/* 3. Empleado */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                        <circle cx="12" cy="7" r="4"/>
                      </svg>
                      Empleado
                    </label>
                    <div className="searchable-select-container">
                      <input
                        ref={empleadoInputRef}
                        type="text"
                        className="select-input"
                        placeholder="Escribe el NIT para buscar..."
                        value={empleadoBusqueda}
                        onChange={(e) => {
                          setEmpleadoBusqueda(e.target.value)
                          recalcDropdownPos()
                          setEmpleadoDropdownVisible(true)
                          if (selectedEmpleado && e.target.value !== selectedEmpleado.nit) {
                            setSelectedEmpleado(null)
                            setTipoDoc('')
                            setTieneReteICA('')
                            setTieneRetencionFuente('')
                            setValorRetencionFuente('')
                          }
                        }}
                        onFocus={() => { recalcDropdownPos(); setEmpleadoDropdownVisible(true) }}
                        onBlur={() => setTimeout(() => setEmpleadoDropdownVisible(false), 150)}
                      />
                      {empleadoDropdownVisible && empleadosFiltrados.length > 0 && createPortal(
                        <ul
                          className="searchable-dropdown"
                          style={{ top: dropdownPortalPos.top, left: dropdownPortalPos.left, width: dropdownPortalPos.width }}
                        >
                          {empleadosFiltrados.map((emp) => (
                            <li
                              key={emp.nit}
                              className={`searchable-dropdown-item${selectedEmpleado?.nit === emp.nit ? ' active' : ''}`}
                              onMouseDown={(e) => e.preventDefault()}
                              onClick={() => handleSeleccionarEmpleado(emp)}
                            >
                              <span className="dropdown-nit">{emp.nit}</span>
                              <span className="dropdown-nombre">{emp.nombreCompleto}</span>
                            </li>
                          ))}
                        </ul>,
                        document.body
                      )}
                    </div>
                    {selectedEmpleado && (
                      <>
                        <div className="empleado-info-panel" style={{marginTop: '0.75rem'}}>
                          <div className="reteica-info-title">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
                              <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
                            </svg>
                            Datos del empleado seleccionado
                            <span className="empleado-info-hint">Puedes editar los campos</span>
                          </div>
                          <div className="empleado-info-grid">
                            <div className="reteica-info-item">
                              <span className="reteica-info-label">NIT / Identificación</span>
                              <input className="empleado-info-input" value={selectedEmpleado.nit} onChange={e => setSelectedEmpleado({...selectedEmpleado, nit: e.target.value})} />
                            </div>
                            <div className="reteica-info-item">
                              <span className="reteica-info-label">Nombre Completo</span>
                              <input className="empleado-info-input" value={selectedEmpleado.nombreCompleto} onChange={e => setSelectedEmpleado({...selectedEmpleado, nombreCompleto: e.target.value})} />
                            </div>
                            <div className="reteica-info-item">
                              <span className="reteica-info-label">Teléfono</span>
                              <input className="empleado-info-input" value={selectedEmpleado.telefono} onChange={e => setSelectedEmpleado({...selectedEmpleado, telefono: e.target.value})} />
                            </div>
                            <div className="reteica-info-item">
                              <span className="reteica-info-label">Dirección</span>
                              <input className="empleado-info-input" value={selectedEmpleado.direccion} onChange={e => setSelectedEmpleado({...selectedEmpleado, direccion: e.target.value})} />
                            </div>
                            <div className="reteica-info-item">
                              <span className="reteica-info-label">Departamento</span>
                              <input className="empleado-info-input" value={selectedEmpleado.departamento} onChange={e => setSelectedEmpleado({...selectedEmpleado, departamento: e.target.value})} />
                            </div>
                          </div>
                        </div>
                        {/* Tipo de Documento */}
                        <div className="form-group" style={{marginTop: '0.75rem'}}>
                          <label className="label">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <rect x="2" y="5" width="20" height="14" rx="2"/>
                              <line x1="2" y1="10" x2="22" y2="10"/>
                            </svg>
                            Tipo de Documento
                          </label>
                          <select
                            className="select-input"
                            value={tipoDoc}
                            onChange={(e) => setTipoDoc(e.target.value)}
                          >
                            <option value="">— Selecciona —</option>
                            <option value="NIT">NIT</option>
                            <option value="CC">CC</option>
                          </select>
                        </div>
                      </>
                    )}
                  </div>

                  {/* 4. Valor */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <line x1="12" y1="1" x2="12" y2="23"/>
                        <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>
                      </svg>
                      Valor
                    </label>
                    <input
                      type="text"
                      className="select-input"
                      placeholder="Ej: $1.500.000"
                      value={valor}
                      onChange={(e) => setValor(e.target.value)}
                    />
                  </div>

                  {/* 5. Concepto */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <line x1="17" y1="10" x2="3" y2="10"/>
                        <line x1="21" y1="6" x2="3" y2="6"/>
                        <line x1="21" y1="14" x2="3" y2="14"/>
                        <line x1="13" y1="18" x2="3" y2="18"/>
                      </svg>
                      Concepto
                    </label>
                    <textarea
                      className="select-input concepto-textarea"
                      placeholder="Describe el concepto de la carta..."
                      value={concepto}
                      onChange={(e) => setConcepto(e.target.value)}
                      rows={3}
                    />
                  </div>

                  {/* 6. Código de Actividad Económica */}
                  <div className="form-group reteica-selector">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M9 11l3 3L22 4"/>
                        <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
                      </svg>
                      Código de Actividad Económica
                      <span className={`badge-source ${reteICASource === 'usuario' ? 'badge-source--usuario' : ''}`}>
                        {reteICASource === 'usuario' ? 'Base actualizada' : 'Base integrada'}
                      </span>
                    </label>
                    <select
                      className="select-input"
                      value={selectedCodigo}
                      onChange={handleCodigoChange}
                      disabled={reteICAData.length === 0}
                    >
                      <option value="">
                        {reteICAData.length === 0 ? 'Cargando...' : '— Selecciona un código —'}
                      </option>
                      {reteICAData.map((item) => (
                        <option key={item.codigo} value={item.codigo}>
                          {item.codigo}
                        </option>
                      ))}
                    </select>
                    {selectedRow && (
                      <div className="reteica-info-panel">
                        <div className="reteica-info-title">
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <circle cx="12" cy="12" r="10"/>
                            <line x1="12" y1="16" x2="12" y2="12"/>
                            <line x1="12" y1="8" x2="12.01" y2="8"/>
                          </svg>
                          Información del código seleccionado
                        </div>
                        <div className="reteica-info-grid">
                          <div className="reteica-info-item">
                            <span className="reteica-info-label">Código</span>
                            <span className="reteica-info-value">{selectedRow.codigo}</span>
                          </div>
                          <div className="reteica-info-item">
                            <span className="reteica-info-label">Descripción</span>
                            <span className="reteica-info-value">{selectedRow.descripcion}</span>
                          </div>
                          <div className="reteica-info-item">
                            <span className="reteica-info-label">Tarifa por Mil</span>
                            <span className="reteica-info-value reteica-tarifa">{selectedRow.tarifaPorMil}</span>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* 7. ¿Tiene Retención de ICA? */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                      </svg>
                      ¿Tiene Retención de ICA?
                    </label>
                    <div className="toggle-group">
                      <button className={`toggle-btn${tieneReteICA === 'si' ? ' toggle-btn--active-yes' : ''}`} onClick={() => setTieneReteICA('si')} type="button">Sí</button>
                      <button className={`toggle-btn${tieneReteICA === 'no' ? ' toggle-btn--active-no' : ''}`} onClick={() => setTieneReteICA('no')} type="button">No</button>
                    </div>
                  </div>

                  {/* 8. ¿Tiene Retención en la Fuente? */}
                  <div className="form-group">
                    <label className="label">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                      </svg>
                      ¿Tiene Retención en la Fuente?
                    </label>
                    <div className="toggle-group">
                      <button className={`toggle-btn${tieneRetencionFuente === 'si' ? ' toggle-btn--active-yes' : ''}`} onClick={() => setTieneRetencionFuente('si')} type="button">Sí</button>
                      <button className={`toggle-btn${tieneRetencionFuente === 'no' ? ' toggle-btn--active-no' : ''}`} onClick={() => setTieneRetencionFuente('no')} type="button">No</button>
                    </div>
                    {tieneRetencionFuente === 'si' && (
                      <input
                        type="text"
                        className="select-input retencion-valor-input"
                        placeholder="Ingresa el valor de Retención en la Fuente"
                        value={valorRetencionFuente}
                        onChange={(e) => setValorRetencionFuente(e.target.value)}
                      />
                    )}
                  </div>

                  {/* Botón Generar Carta */}
                  <div className="section-divider" style={{marginTop: '1rem'}} />
                  <div className="generar-carta-wrap">
                    <div className="generar-btns-row">
                      <button
                        className="btn-generar"
                        onClick={generateCarta}
                        disabled={generando || generandoPDF || !selectedEmpleado || !fechaProceso}
                      >
                        {generando ? (
                          <>
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="spin">
                              <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                            </svg>
                            Generando...
                          </>
                        ) : (
                          <>
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                              <polyline points="14 2 14 8 20 8"/>
                              <line x1="12" y1="18" x2="12" y2="12"/>
                              <line x1="9" y1="15" x2="15" y2="15"/>
                            </svg>
                            Generar Word
                          </>
                        )}
                      </button>
                      <button
                        className="btn-generar btn-generar--pdf"
                        onClick={generatePDF}
                        disabled={generando || generandoPDF || !selectedEmpleado || !fechaProceso}
                      >
                        {generandoPDF ? (
                          <>
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="spin">
                              <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                            </svg>
                            Generando PDF...
                          </>
                        ) : (
                          <>
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                              <polyline points="14 2 14 8 20 8"/>
                              <path d="M9 13h6M9 17h3"/>
                            </svg>
                            Generar PDF
                          </>
                        )}
                      </button>
                    </div>
                    {(!selectedEmpleado || !fechaProceso) && (
                      <p className="generar-hint">Selecciona un empleado y una fecha para habilitar la generación.</p>
                    )}
                  </div>

                </div>
              </div>
            </section>}

          </div>{/* /tab-content */}

        </div>
      </main>

      {/* Modal: Base de Empleados requerida */}
      {showBlockedModal && (
        <div className="modal-overlay" onClick={() => setShowBlockedModal(false)}>
          <div className="modal-box" onClick={(e) => e.stopPropagation()}>
            <div className="modal-icon">
              <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
                <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
              </svg>
            </div>
            <h3>Base de Empleados requerida</h3>
            <p>Para acceder al editor debes cargar primero el archivo Excel de empleados en la sección <strong>Carga de Archivos</strong>.</p>
            <div className="modal-actions">
              <button
                className="modal-btn-ok"
                onClick={() => { setShowBlockedModal(false); setActiveTab('archivos') }}
              >
                Ir a Carga de Archivos
              </button>
              <button className="modal-btn-cancel" onClick={() => setShowBlockedModal(false)}>
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>
    </div>
  )
}

export default App
