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
