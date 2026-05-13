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

    // Assume no duplicate for the current main file
    for (const row of mainMapped) {
        const key = createTransactionKey(row)
        /* if (!mainSeen.has(key)) {} */
        seen.add(key)
        unique.push(row)
    }

    // Anta at det ikke er duplikater i srbank eller duett linjene
    for (const row of [...srbankMapped, ...duettMapped]) {
        const key = createTransactionKey(row)
        if (!seen.has(key)) {
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
    dato = excelDateToNorwegianSrbank(dato)
    const inn = parseFloat(row['Inn'] || 0)
    const ut = parseFloat(row['Ut'] || 0)
    const netto = inn + ut
    const nettoOut = isNaN(netto) ? '' : netto
    return {
        Dato: dato,
        System: 'SR-bank',
        Inn: row['Inn'] || '',
        Ut: row['Ut'] || '',
        Netto: nettoOut,
        'SR+ DU-': nettoOut,
        Avstemming: '',
        'SR-bank: Type': row['Type'] || '',
        'SR-bank: Fra': row['Fra konto'] || '',
        'SR-bank: Til': row['Til konto'] || '',
        Bilag: '',
        'Duett: Periode': '',
        Beskrivelse: row['Beskrivelse'] || '',
    }
}

function mapDuettToMain(row) {
    let dato = row['Dato'] || ''
    dato = excelDateToNorwegian(dato)
    const debet = parseFloat(row['Debet'] || 0)
    const kredit = parseFloat(row['Kredit'] || 0)
    const netto = debet - kredit
    const nettoOut = isNaN(netto) ? '' : netto
    const srDuOut = isNaN(netto) ? '' : -netto
    return {
        Dato: dato,
        System: 'Duett',
        Inn: row['Debet'] || '',
        Ut: row['Kredit'] || '',
        Netto: nettoOut,
        'SR+ DU-': srDuOut,
        Avstemming: '',
        'SR-bank: Type': '',
        'SR-bank: Fra': '',
        'SR-bank: Til': '',
        Bilag: row['Bilag'] || '',
        'Duett: Periode': row['Merknad'] || '',
        Beskrivelse: '',
    }
}
