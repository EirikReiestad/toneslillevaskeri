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
