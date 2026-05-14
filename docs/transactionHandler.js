// Create a unique identifier for a transaction
function createTransactionKey(row) {
    const date = row['Dato'] || ''
    const system = row['System'] || ''
    const netto = parseFloat(row['Netto'] || 0).toFixed(2)
    const bilag = row['Bilag'] || ''
    const beskrivelse = row['Beskrivelse'] || ''

    return `${date}|${system}|${netto}|${bilag}|${beskrivelse}`.toLowerCase()
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

function matchAllTransactions(mainRows, srbankRows, duettRows) {
    const all = mergeRows(mainRows, srbankRows, duettRows)
    total_length = all.length

    // Sort by date so first-match pairing finds nearest in time
    all.sort((a, b) => {
        const da = parseNorwegianDate(a['Dato'])
        const db = parseNorwegianDate(b['Dato'])
        return da - db
    })

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
                    all[i]['Dato'].split('/').reverse().join('-')
                )
                const d2 = new Date(
                    all[j]['Dato'].split('/').reverse().join('-')
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
        const netto = parseFloat(all[i]['Netto'] || 0)
        const inn = all[i]['Inn']
        const ut = all[i]['Ut']
        if (inn !== '' && inn != null) {
            all[i]['Inn'] = Math.abs(parseFloat(inn) || 0)
            continue
        }
        if (ut !== '' && ut != null) {
            all[i]['Ut'] = Math.abs(parseFloat(ut) || 0)
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
