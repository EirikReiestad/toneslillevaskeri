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
