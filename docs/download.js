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
                const [d, m, y] = str.split('/')
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
