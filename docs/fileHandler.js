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
