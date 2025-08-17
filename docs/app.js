console.log(`Version 0.0.8`)

handleFileInput(srbankInput, (data) => {
    srbankData = data
    srbankData.map((row) => {
        const newRow = row
        newRow['Dato'] = newRow['Dato'].replace('.', '/')
        return newRow
    })
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
