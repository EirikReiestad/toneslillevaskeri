// File input elements
const srbankInput = document.getElementById('srbank-upload')
const duettInput = document.getElementById('duett-upload')
const dataPreview = document.getElementById('data-preview')
const inputView = document.getElementById('input-view')
const reviewView = document.getElementById('review-view')
const navInput = document.getElementById('nav-input')
const navReview = document.getElementById('nav-review')

// Add main file input
const mainInput = document.getElementById('main-upload')

let srbankData = null
let duettData = null
let mainData = null
let currentSortMode = 'allByDate' // 'newOnTop' or 'allByDate'

// Save stats to not calculate multiple times
let total_length = 0

// Excel serial date to Norwegian string
function excelDateToNorwegian(serial) {
    if (typeof serial !== 'number') return serial
    const utc_days = Math.floor(serial - 25569)
    const utc_value = utc_days * 86400
    const date_info = new Date(utc_value * 1000)
    const day = String(date_info.getUTCDate()).padStart(2, '0')
    const month = String(date_info.getUTCMonth() + 1).padStart(2, '0')
    const year = date_info.getUTCFullYear()
    return `${day}.${month}.${year}`
}

// Trim column headers in Excel data
function trimColumnHeaders(jsonData) {
    if (!Array.isArray(jsonData) || jsonData.length === 0) return jsonData

    return jsonData.map((row) => {
        const trimmedRow = {}
        Object.keys(row).forEach((key) => {
            const trimmedKey = key.trim()
            trimmedRow[trimmedKey] = row[key]
        })
        return trimmedRow
    })
}

function renderPreview() {
    dataPreview.innerHTML = ''
    // Always render all three preview containers for consistent layout
    const previews = [
        { data: mainData, label: 'Main (Bankavstemming)', id: 'main-preview' },
        { data: srbankData, label: 'SRBank', id: 'srbank-preview' },
        { data: duettData, label: 'Duett', id: 'duett-preview' },
    ]
    previews.forEach((preview) => {
        const container = createTablePreview(preview.data || [], preview.label)
        container.id = preview.id
        dataPreview.appendChild(container)
    })

    // Update stats after rendering preview
    updateStats()
}

function createTablePreview(data, label) {
    const container = document.createElement('div')
    container.className = 'flex-1'
    const title = document.createElement('h3')
    title.className = 'font-semibold mb-2'
    // Norwegian preview titles
    let norskLabel = label
    if (label === 'Main (Bankavstemming)')
        norskLabel = 'Hovedfil (Bankavstemming)'
    if (label === 'SRBank') norskLabel = 'SR-Bank'
    if (label === 'Duett') norskLabel = 'Duett'
    title.textContent = norskLabel + ' forhåndsvisning'
    container.appendChild(title)
    // Add scrollable wrapper
    const scrollWrap = document.createElement('div')
    scrollWrap.className =
        'overflow-x-auto overflow-y-auto max-h-48 bg-gray-50 rounded shadow-inner w-full'
    const table = document.createElement('table')
    table.className = 'min-w-full text-xs border border-gray-200'
    if (data.length === 0) {
        table.innerHTML = '<tr><td class="p-2">No data</td></tr>'
        scrollWrap.appendChild(table)
        container.appendChild(scrollWrap)
        return container
    }
    // Header
    const thead = document.createElement('thead')
    thead.className = 'sticky top-0 bg-blue-50 z-10'
    const headerRow = document.createElement('tr')
    Object.keys(data[0]).forEach((key) => {
        const th = document.createElement('th')
        th.className = 'p-2 border-b text-left'
        th.textContent = key
        headerRow.appendChild(th)
    })
    thead.appendChild(headerRow)
    table.appendChild(thead)
    // Body
    const tbody = document.createElement('tbody')
    // For Duett, skip first two rows in preview
    let previewRows = data
    if (label === 'Duett') previewRows = data.slice(2)
    // For SRBank, skip empty rows
    if (label === 'SRBank') {
        previewRows = data.filter((row) =>
            Object.values(row).some((val) => val !== '')
        )
    }
    previewRows.forEach((row) => {
        const tr = document.createElement('tr')
        tr._rowData = row // Store the row data with the table row element
        Object.entries(row).forEach(([key, val]) => {
            const td = document.createElement('td')
            td.className = 'p-2 border-b'
            td.textContent = key === 'Dato' ? excelDateToNorwegian(val) : val
            tr.appendChild(td)
        })
        tbody.appendChild(tr)
    })
    table.appendChild(tbody)
    scrollWrap.appendChild(table)
    container.appendChild(scrollWrap)
    return container
}

function handleFileInput(input, setData) {
    input.addEventListener('change', (e) => {
        const file = e.target.files[0]
        if (!file) return
        if (input === mainInput) {
            setData(null, file) // Custom handler for main
            return
        }
        const reader = new FileReader()
        reader.onload = (evt) => {
            const data = new Uint8Array(evt.target.result)
            const workbook = XLSX.read(data, { type: 'array' })
            // Use the first sheet
            const sheetName = workbook.SheetNames[0]
            const sheet = workbook.Sheets[sheetName]
            const json = XLSX.utils.sheet_to_json(sheet, { defval: '' })
            setData(trimColumnHeaders(json))
            renderPreview()
            updateStats()
        }
        reader.readAsArrayBuffer(file)
    })
}

handleFileInput(srbankInput, (data) => {
    srbankData = data
    renderPreview()
    updateStats()
})
handleFileInput(duettInput, (data) => {
    // Skip first two rows (headers/irrelevant)
    duettData = Array.isArray(data) ? data.slice(2) : data
    renderPreview()
    updateStats()
})
handleFileInput(mainInput, (data, file) => {
    // Custom handler for main file: fetch 'Bankavstemming' sheet if present
    const reader = new FileReader()
    reader.onload = (evt) => {
        const dataArr = new Uint8Array(evt.target.result)
        const workbook = XLSX.read(dataArr, { type: 'array' })
        let sheetName =
            workbook.SheetNames.find(
                (n) => n.toLowerCase() === 'bankavstemming'
            ) || workbook.SheetNames[0]
        const sheet = workbook.Sheets[sheetName]
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' })
        mainData = trimColumnHeaders(json)
        // Store main workbook for download
        window.mainWorkbook = workbook
        renderPreview()
        updateStats()
    }
    reader.readAsArrayBuffer(file)
})

// Navigation
navInput.addEventListener('click', (e) => {
    e.preventDefault()
    inputView.style.display = ''
    reviewView.style.display = 'none'
    navInput.classList.add('text-gray-800', 'border-blue-600')
    navInput.classList.remove('text-gray-600', 'border-transparent')
    navReview.classList.remove('text-gray-800', 'border-blue-600')
    navReview.classList.add('text-gray-600', 'border-transparent')
})

// --- Mapping helpers ---
const MAIN_COLUMNS = [
    'Dato',
    'System',
    'Inn',
    'Ut',
    'Netto',
    'SR+ DU-',
    'Avstemming',
    'SR-bank: Type',
    'SR-bank: Fra',
    'SR-bank: Til',
    'Bilag',
    'Duett: Periode',
    'Beskrivelse',
]

function isValidDateString(dateStr) {
    // Accepts DD.MM.YYYY, must not be empty and must not contain any letters
    if (!dateStr || typeof dateStr !== 'string') return false
    if (/[a-zA-Z]/.test(dateStr)) return false
    // Optionally, check for format: two digits, dot, two digits, dot, four digits
    if (!/^\d{2}\.\d{2}\.\d{4}$/.test(dateStr)) return false
    return true
}

// Create a unique identifier for a transaction
function createTransactionKey(row) {
    // Use date, system, amount, and description to create a unique key
    const date = row['Dato'] || ''
    const system = row['System'] || ''
    const netto = parseFloat(row['Netto'] || 0).toFixed(2)
    const bilag = row['Bilag'] || ''

    return `${date}|${system}|${netto}|${bilag}`.toLowerCase()
}

// Remove duplicate transactions from a list
function removeDuplicates(rows) {
    const seen = new Set()
    const unique = []

    for (const row of rows) {
        const key = createTransactionKey(row)
        if (!seen.has(key)) {
            seen.add(key)
            unique.push(row)
        }
    }

    return unique
}

function mergeRows(mainRows, srbankRows, duettRows) {
    // Only keep rows with a valid date and remove duplicates
    const mainMapped = mainRows
        .map(mapMainRow)
        .filter((row) => isValidDateString(row['Dato']))
    const srbankMapped = srbankRows
        .map(mapSRBankToMain)
        .filter((row) => isValidDateString(row['Dato']))
    const duettMapped = duettRows
        .map(mapDuettToMain)
        .filter((row) => isValidDateString(row['Dato']))

    const seen = new Set()
    const unique = []

    for (const row of mainMapped) {
        const key = createTransactionKey(row)
        if (!seen.has(key)) {
            seen.add(key)
            unique.push(row)
        }
    }
    for (const row of [...srbankMapped, ...duettMapped]) {
        const key = createTransactionKey(row)
        if (!seen.has(key)) {
            seen.add(key)
            unique.push(row)
            continue
        }
    }

    return unique
}

function mapMainRow(row) {
    // Ensure all columns exist
    const out = {}
    MAIN_COLUMNS.forEach((col) => {
        if (col === 'Dato') {
            out[col] = excelDateToNorwegian(row[col])
        } else {
            out[col] = row[col] || ''
        }
    })
    return out
}

function mapSRBankToMain(row) {
    let dato = row['Dato'] || ''
    dato = excelDateToNorwegian(dato)
    return {
        Dato: dato,
        System: 'SR-bank',
        Inn: row['Inn'] || '',
        Ut: row['Ut'] || '',
        Netto: parseFloat(row['Inn'] || 0) + parseFloat(row['Ut'] || 0) || '',
        'SR+ DU-':
            parseFloat(row['Inn'] || 0) + parseFloat(row['Ut'] || 0) || '',
        Avstemming: '',
        'SR-bank: Type': row['Type'] || '',
        'SR-bank: Fra': row['Fra'] || '',
        'SR-bank: Til': row['Til'] || '',
        Bilag: '',
        'Duett: Periode': '',
        Beskrivelse: row['Beskrivelse'] || '',
    }
}

function mapDuettToMain(row) {
    let dato = row['Dato'] || ''
    dato = excelDateToNorwegian(dato)
    return {
        Dato: dato,
        System: 'Duett',
        Inn: row['Debet'] || '',
        Ut: row['Kredit'] || '',
        Netto:
            parseFloat(row['Debet'] || 0) - parseFloat(row['Kredit'] || 0) ||
            '',
        'SR+ DU-': -(
            parseFloat(row['Debet'] || 0) - parseFloat(row['Kredit'] || 0) || ''
        ),
        Avstemming: '',
        'SR-bank: Type': '',
        'SR-bank: Fra': '',
        'SR-bank: Til': '',
        Bilag: row['Bilag'] || '',
        'Duett: Periode': row['Merknad'] || '',
        Beskrivelse: '',
    }
}

function matchAllTransactions(mainRows, srbankRows, duettRows) {
    const all = mergeRows(mainRows, srbankRows, duettRows)
    total_length = all.length

    all.sort((a, b) =>
        a['Dato'] > b['Dato'] ? 1 : a['Dato'] < b['Dato'] ? -1 : 0
    )
    let maxANum = 0
    all.forEach((row) => {
        if (row['Avstemming'] && /^A\d+$/.test(row['Avstemming'])) {
            const num = parseInt(row['Avstemming'].slice(1))
            if (num > maxANum) maxANum = num
        }
        row._oldPair = !!row['Avstemming']
        row._newlyMatched = false
        row._matched = false
    })
    let tagCounter = maxANum + 1
    const used = new Array(all.length).fill(false)
    for (let i = 0; i < all.length; ++i) {
        if (used[i]) continue
        if (all[i]['Avstemming']) continue // skip already matched
        for (let j = i + 1; j < all.length; ++j) {
            if (used[j]) continue
            if (all[j]['Avstemming']) continue
            if (all[i]['System'] === all[j]['System']) continue
            const n1 = parseFloat(all[i]['Netto'] || 0)
            const n2 = parseFloat(all[j]['Netto'] || 0)
            if (Math.abs(n1 - n2) < 0.01) {
                const d1 = new Date(
                    all[i]['Dato'].split('.').reverse().join('-')
                )
                const d2 = new Date(
                    all[j]['Dato'].split('.').reverse().join('-')
                )
                const diffDays = Math.abs((d1 - d2) / (1000 * 60 * 60 * 24))
                if (diffDays <= 3) {
                    const tag = `A${tagCounter++}`
                    all[i]['Avstemming'] = tag
                    all[j]['Avstemming'] = tag
                    all[i]._newlyMatched = true
                    all[j]._newlyMatched = true
                    all[i]._matched = true
                    all[j]._matched = true
                    used[i] = true
                    used[j] = true
                    break
                }
            }
        }
    }
    for (let i = 0; i < all.length; ++i) {
        // Utter trash code:)
        const netto = parseFloat(all[i]['Netto'] || 0)
        if (all[i]['Inn'] || all[i]['Inn'] !== '') {
            all[i]['Inn'] = Math.abs(all[i]['Inn'])
            continue
        }
        if (all[i]['Ut'] || all[i]['Ut'] !== '') {
            all[i]['Ut'] = Math.abs(all[i]['Ut'])
            continue
        }
        if (netto > 0) {
            all[i]['Inn'] = Math.abs(netto)
        } else if (netto < 0) {
            all[i]['Ut'] = Math.abs(netto)
        }
    }
    return all
}

// --- Review Table ---
function parseNorwegianDate(dateStr) {
    // Converts DD.MM.YYYY to YYYY-MM-DD for sorting
    if (!dateStr || typeof dateStr !== 'string') return ''
    const parts = dateStr.split('.')
    if (parts.length !== 3) return ''
    return new Date(parts[2], parts[1] - 1, parts[0])
    // return `${parts[2]}-${parts[1]}-${parts[0]}`
}

function renderReviewTable(combined) {
    reviewView.innerHTML = ''
    if (!combined || combined.length === 0) {
        reviewView.innerHTML =
            '<div class="text-gray-500">Ingen data å gjennomgå.</div>'
        return
    }
    // Outer wrapper for margin
    const outerWrap = document.createElement('div')
    outerWrap.className = 'px-4 md:px-12'
    // Button bar (sort + download, right-aligned)
    const btnBar = document.createElement('div')
    btnBar.className = 'flex items-center mb-2 justify-end gap-2' // right-aligned, with gap
    // Sort button
    const sortBtn = document.createElement('button')
    sortBtn.className =
        'px-2 py-2 bg-gray-200 rounded hover:bg-gray-300 flex items-center justify-center'
    sortBtn.title =
        currentSortMode === 'newOnTop'
            ? 'Sorter: Nye matcher øverst (trykk for kun dato)'
            : 'Sorter: Kun dato (trykk for nye matcher øverst)'
    sortBtn.innerHTML =
        currentSortMode === 'newOnTop'
            ? '<span style="font-size:1.2em;">&#128200;</span>' // bar chart icon
            : '<span style="font-size:1.2em;">&#128197;</span>' // calendar icon
    sortBtn.onclick = () => {
        currentSortMode =
            currentSortMode === 'newOnTop' ? 'allByDate' : 'newOnTop'
        renderReviewTable(combined)
    }
    btnBar.appendChild(sortBtn)
    // Excel download icon button
    const excelBtn = document.createElement('button')
    excelBtn.className =
        'px-2 py-2 bg-green-600 text-white rounded hover:bg-green-700 flex items-center justify-center'
    excelBtn.title = 'Last ned som Excel'
    excelBtn.innerHTML =
        '<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" width="20" height="20"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M7 10l5 5m0 0l5-5m-5 5V4" /></svg>'
    excelBtn.onclick = () => downloadExcelResults(combined)
    btnBar.appendChild(excelBtn)
    outerWrap.appendChild(btnBar)
    // Add scrollable wrapper for the table
    const scrollWrap = document.createElement('div')
    scrollWrap.className =
        'w-full overflow-x-auto overflow-y-auto max-h-[70vh] bg-white rounded shadow-inner'
    const table = document.createElement('table')
    table.className = 'min-w-full text-xs border border-gray-300'
    const thead = document.createElement('thead')
    thead.className = 'sticky top-0 bg-blue-50 z-10'
    const headerRow = document.createElement('tr')
    // Add row number header
    const thRowNum = document.createElement('th')
    thRowNum.className = 'p-2 border-b text-left'
    thRowNum.textContent = '#'
    headerRow.appendChild(thRowNum)
    MAIN_COLUMNS.forEach((h) => {
        const th = document.createElement('th')
        th.className = 'p-2 border-b text-left'
        th.textContent = h
        headerRow.appendChild(th)
    })
    thead.appendChild(headerRow)
    table.appendChild(thead)
    const tbody = document.createElement('tbody')
    const pairColors = ['bg-green-200', 'bg-green-100']
    let colorIdx = 0
    let rowNum = 1
    let highlightedRows = new Set() // Track currently highlighted rows

    // Function to highlight matching rows
    function highlightMatchingRows(avstemmingTag) {
        // Clear previous highlights
        highlightedRows.clear()
        const allTableRows = tbody.querySelectorAll('tr')
        allTableRows.forEach((row) => {
            row.classList.remove('bg-yellow-200', 'ring-2', 'ring-yellow-400')
        })

        if (avstemmingTag) {
            // Find and highlight all rows with the same Avstemming tag
            allTableRows.forEach((row) => {
                const rowData = row._rowData // Get the data stored with the row
                if (rowData && rowData['Avstemming'] === avstemmingTag) {
                    row.classList.add(
                        'bg-yellow-200',
                        'ring-2',
                        'ring-yellow-400'
                    )
                }
            })
        }
    }

    if (currentSortMode === 'allByDate') {
        // Show all transactions, sorted by full date
        const allRows = [...combined]
        allRows.sort((a, b) => {
            const da = parseNorwegianDate(a['Dato'])
            const db = parseNorwegianDate(b['Dato'])
            return da - db
            // return da > db ? 1 : da < db ? -1 : 0
        })
        allRows.forEach((row) => {
            const tr = document.createElement('tr')
            tr._rowData = row // Store the row data with the table row element
            if (row._matched) {
                tr.className = pairColors[colorIdx % pairColors.length]
                colorIdx++
            } else if (!row['Avstemming']) {
                // Add subtle background for unmatched rows
                tr.className = 'bg-gray-50'
            }

            // Add click handler
            tr.addEventListener('click', () => {
                highlightMatchingRows(row['Avstemming'])
            })

            // Add cursor pointer if row has Avstemming
            if (row['Avstemming']) {
                tr.style.cursor = 'pointer'
                // tr.title = 'Klikk for å fremheve matchende rader'
            }

            // Row number
            const tdRowNum = document.createElement('td')
            tdRowNum.className = 'p-2 border-b text-right'
            tdRowNum.textContent = rowNum++
            tr.appendChild(tdRowNum)
            MAIN_COLUMNS.forEach((col) => {
                const td = document.createElement('td')
                td.className = 'p-2 border-b'
                td.textContent = row[col] || ''
                tr.appendChild(td)
            })
            tbody.appendChild(tr)
        })
    } else {
        // New matches on top (sorted by date), then others (also sorted by date)
        // Group by Avstemming tag (for new matches only)
        const pairs = {}
        const unmatched = []
        const oldPairs = []
        combined.forEach((row) => {
            if (row._matched) {
                if (!pairs[row['Avstemming']]) pairs[row['Avstemming']] = []
                pairs[row['Avstemming']].push(row)
            } else if (row._oldPair) {
                oldPairs.push(row)
            } else {
                unmatched.push(row)
            }
        })
        const pairTags = Object.keys(pairs).sort((a, b) => {
            const na = parseInt(a.slice(1))
            const nb = parseInt(b.slice(1))
            return nb - na
        })
        let newPairRows = []
        pairTags.forEach((tag) => {
            newPairRows = newPairRows.concat(pairs[tag])
        })
        newPairRows.sort((a, b) => {
            const da = parseNorwegianDate(a['Dato'])
            const db = parseNorwegianDate(b['Dato'])
            return da > db ? 1 : da < db ? -1 : 0
        })
        newPairRows.forEach((row) => {
            const tr = document.createElement('tr')
            tr._rowData = row // Store the row data with the table row element
            tr.className = pairColors[colorIdx % pairColors.length]

            // Add click handler
            tr.addEventListener('click', () => {
                highlightMatchingRows(row['Avstemming'])
            })

            // Add cursor pointer if row has Avstemming
            if (row['Avstemming']) {
                tr.style.cursor = 'pointer'
                // tr.title = 'Klikk for å fremheve matchende rader'
            }

            // Row number
            const tdRowNum = document.createElement('td')
            tdRowNum.className = 'p-2 border-b text-right'
            tdRowNum.textContent = rowNum++
            tr.appendChild(tdRowNum)
            MAIN_COLUMNS.forEach((col) => {
                const td = document.createElement('td')
                td.className = 'p-2 border-b'
                td.textContent = row[col] || ''
                tr.appendChild(td)
            })
            tbody.appendChild(tr)
        })
        // Unmatched and old pairs, sorted by date
        const restRows = [...unmatched, ...oldPairs]
        restRows.sort((a, b) => {
            const da = parseNorwegianDate(a['Dato'])
            const db = parseNorwegianDate(b['Dato'])
            return da > db ? 1 : da < db ? -1 : 0
        })
        restRows.forEach((row) => {
            const tr = document.createElement('tr')
            tr._rowData = row // Store the row data with the table row element

            // Add subtle background for unmatched rows
            if (!row['Avstemming']) {
                tr.className = 'bg-gray-50'
            }

            // Add click handler
            tr.addEventListener('click', () => {
                highlightMatchingRows(row['Avstemming'])
            })

            // Add cursor pointer if row has Avstemming
            if (row['Avstemming']) {
                tr.style.cursor = 'pointer'
                //tr.title = 'Klikk for å fremheve matchende rader'
            }

            // Row number
            const tdRowNum = document.createElement('td')
            tdRowNum.className = 'p-2 border-b text-right'
            tdRowNum.textContent = rowNum++
            tr.appendChild(tdRowNum)
            MAIN_COLUMNS.forEach((col) => {
                const td = document.createElement('td')
                td.className = 'p-2 border-b'
                td.textContent = row[col] || ''
                tr.appendChild(td)
            })
            tbody.appendChild(tr)
        })
    }
    table.appendChild(tbody)
    scrollWrap.appendChild(table)
    outerWrap.appendChild(scrollWrap)
    reviewView.appendChild(outerWrap)
}

function downloadExcelResults(combined) {
    if (!combined || combined.length === 0) {
        alert('Ingen data å laste ned.')
        return
    }

    // Generate timestamp for filename
    const now = new Date()
    const dateStr = now
        .toISOString()
        .slice(0, 19)
        .replace(/:/g, '-')
        .replace('T', '_')
    const filename = `bankavstemming_${dateStr}.xlsx`

    try {
        // Prepare data for Excel
        const headers = MAIN_COLUMNS // No 'Matched' column

        const sortedData = [...combined].sort((a, b) => {
            const parse = (str) => {
                if (!str || typeof str !== 'string') return ''
                const [d, m, y] = str.split('.')
                return new Date(`${y}-${m}-${d}`)
            }
            return parse(a['Dato']) - parse(b['Dato'])
        })

        const data = sortedData.map((row) => {
            const out = {}
            headers.forEach((h) => {
                out[h] = row[h] || ''
            })
            return out
        })

        const ws = XLSX.utils.json_to_sheet(data, { header: headers })
        const wb = XLSX.utils.book_new()
        XLSX.utils.book_append_sheet(wb, ws, 'Bankavstemming')

        // Try the standard download method first
        try {
            XLSX.writeFile(wb, filename)
        } catch (downloadError) {
            console.warn(
                'Standard download failed, trying alternative method:',
                downloadError
            )

            // Alternative download method for problematic browsers
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
            const blob = new Blob([wbout], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            })

            // Create download link
            const url = window.URL.createObjectURL(blob)
            const a = document.createElement('a')
            a.href = url
            a.download = filename
            document.body.appendChild(a)
            a.click()

            // Cleanup
            setTimeout(() => {
                document.body.removeChild(a)
                window.URL.revokeObjectURL(url)
            }, 100)
        }

        console.log('Download completed successfully')
    } catch (error) {
        console.error('Download failed:', error)
        alert(
            'Nedlasting feilet. Vennligst prøv igjen eller sjekk at filen ikke er åpen i Excel.'
        )
    }
}

// --- Hook up Review Tab ---
navReview.addEventListener('click', (e) => {
    e.preventDefault()
    inputView.style.display = 'none'
    reviewView.style.display = ''
    navReview.classList.add('text-gray-800', 'border-blue-600')
    navReview.classList.remove('text-gray-600', 'border-transparent')
    navInput.classList.remove('text-gray-800', 'border-blue-600')
    navInput.classList.add('text-gray-600', 'border-transparent')
    // Only run if all files are loaded
    if (mainData && srbankData && duettData) {
        const combined = matchAllTransactions(mainData, srbankData, duettData)
        renderReviewTable(combined)
    } else {
        reviewView.innerHTML =
            '<div class="text-gray-500 text-center mt-12">Vennligst last opp alle tre filene for å gjennomgå matcher.</div>'
    }
})

// Load sample files for debugging
async function loadSampleFiles() {
    try {
        // Load main file
        const mainResponse = await fetch('sample_main.xlsx')
        const mainBuffer = await mainResponse.arrayBuffer()
        const mainDataArr = new Uint8Array(mainBuffer)
        const mainWorkbook = XLSX.read(mainDataArr, { type: 'array' })
        let mainSheetName =
            mainWorkbook.SheetNames.find(
                (n) => n.toLowerCase() === 'bankavstemming'
            ) || mainWorkbook.SheetNames[0]
        const mainSheet = mainWorkbook.Sheets[mainSheetName]
        mainData = trimColumnHeaders(
            XLSX.utils.sheet_to_json(mainSheet, { defval: '' })
        )

        // Load SRBank file
        const srbankResponse = await fetch('sample_srbank.xlsx')
        const srbankBuffer = await srbankResponse.arrayBuffer()
        const srbankDataArr = new Uint8Array(srbankBuffer)
        const srbankWorkbook = XLSX.read(srbankDataArr, { type: 'array' })
        const srbankSheet = srbankWorkbook.Sheets[srbankWorkbook.SheetNames[0]]
        srbankData = trimColumnHeaders(
            XLSX.utils.sheet_to_json(srbankSheet, { defval: '' })
        )

        // Load Duett file
        const duettResponse = await fetch('sample_duett.xlsx')
        const duettBuffer = await duettResponse.arrayBuffer()
        const duettDataArr = new Uint8Array(duettBuffer)
        const duettWorkbook = XLSX.read(duettDataArr, { type: 'array' })
        const duettSheet = duettWorkbook.Sheets[duettWorkbook.SheetNames[0]]
        const duettJson = trimColumnHeaders(
            XLSX.utils.sheet_to_json(duettSheet, { defval: '' })
        )
        duettData = Array.isArray(duettJson) ? duettJson.slice(2) : duettJson

        // Store main workbook for download
        window.mainWorkbook = mainWorkbook

        // Render preview
        renderPreview()

        // Update stats
        updateStats()
    } catch (error) {
        console.error('Error loading sample files:', error)
    }
}

// Load sample files when page loads
// document.addEventListener('DOMContentLoaded', loadSampleFiles)

// Stats calculation and display functions
function updateStats() {
    // Count only valid rows (with valid dates)
    const mainCount = mainData
        ? mainData.filter((row) =>
              isValidDateString(excelDateToNorwegian(row['Dato']))
          ).length
        : 0
    const srbankCount = srbankData
        ? srbankData.filter((row) =>
              isValidDateString(excelDateToNorwegian(row['Dato']))
          ).length
        : 0
    const duettCount = duettData
        ? duettData.filter((row) =>
              isValidDateString(excelDateToNorwegian(row['Dato']))
          ).length
        : 0

    // Update file counts
    document.getElementById('main-count').textContent = mainCount
    document.getElementById('srbank-count').textContent = srbankCount
    document.getElementById('duett-count').textContent = duettCount

    // Calculate expected total (main + new transactions)
    const expectedTotal = total_length
    document.getElementById('expected-total').textContent = expectedTotal

    // Calculate matching statistics if all files are loaded
    if (mainData && srbankData && duettData) {
        const combined = matchAllTransactions(mainData, srbankData, duettData)

        // Count existing matches (rows with Avstemming that are not new)
        const existingMatches = combined.filter(
            (row) => row['Avstemming'] && !row._matched
        ).length
        const newMatches = combined.filter((row) => row._matched).length

        // Calculate percentages relative to expected total (main + new matches)
        const existingPercentage =
            expectedTotal > 0
                ? Math.round((existingMatches / expectedTotal) * 100)
                : 0
        const newPercentage =
            expectedTotal > 0
                ? Math.round((newMatches / expectedTotal) * 100)
                : 0

        // Count total matches (existing + new)
        const totalMatches = existingMatches + newMatches
        const totalPercentage =
            expectedTotal > 0
                ? Math.round((totalMatches / expectedTotal) * 100)
                : 0

        // Update stats
        document.getElementById('existing-matches').textContent =
            existingPercentage
        document.getElementById('new-matches').textContent = newPercentage
        document.getElementById('total-matches').textContent = totalPercentage
    } else {
        // Reset stats if not all files loaded
        document.getElementById('existing-matches').textContent = '-'
        document.getElementById('new-matches').textContent = '-'
        document.getElementById('total-matches').textContent = '-'
    }
}
