function isValidDateString(dateStr) {
    // Accepts YYYY/MM/DD, must not be empty and must not contain any letters
    if (!dateStr || typeof dateStr !== 'string') {
        return false
    }
    if (/[a-zA-Z]/.test(dateStr)) return false
    // Optionally, check for format: two digits, dot, two digits, dot, four digits
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) return false
    return true
}

// --- Review Table ---
function parseNorwegianDate(dateStr) {
    // Converts DD/MM/YYYY to YYYY-MM-DD for sorting
    if (!dateStr || typeof dateStr !== 'string') return ''
    const parts = dateStr.split('/')
    if (parts.length !== 3) return ''
    return new Date(parts[2], parts[1] - 1, parts[0])
    // return `${parts[2]}-${parts[1]}-${parts[0]}`
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

// Excel serial date to Norwegian string
function excelDateToNorwegian(date) {
    if (typeof date === 'number') return serialDateToDate(date)
    if (typeof date !== 'string' || date === '') return ''
    const splitDate = date.split('/').join(',').split('.').join(',').split(',')
    if (splitDate.length !== 3) return ''
    const [day, month, year] = splitDate
    return `${day}/${month}/${year}`
}

const excelDateToNorwegianSrbank = excelDateToNorwegian

function serialDateToDate(serial) {
    if (typeof serial !== 'number') return serial
    const utc_days = Math.floor(serial - 25569)
    const utc_value = utc_days * 86400
    const date_info = new Date(utc_value * 1000)
    const day = String(date_info.getUTCDate()).padStart(2, '0')
    const month = String(date_info.getUTCMonth() + 1).padStart(2, '0')
    const year = date_info.getUTCFullYear()
    return `${day}/${month}/${year}`
}
