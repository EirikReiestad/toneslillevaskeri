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
