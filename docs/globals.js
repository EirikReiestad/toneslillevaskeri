// --- Mapping helpers ---
const MAIN_COLUMNS = [
    'Dato',
    'System',
    'Inn',
    'Ut',
    'Netto',
    'SR+ DU-',
    'Avstemming',
    'SR-bank: Type',
    'SR-bank: Fra',
    'SR-bank: Til',
    'Bilag',
    'Duett: Periode',
    'Beskrivelse',
]
// File input elements
const srbankInput = document.getElementById('srbank-upload')
const duettInput = document.getElementById('duett-upload')
const dataPreview = document.getElementById('data-preview')
const inputView = document.getElementById('input-view')
const reviewView = document.getElementById('review-view')
const navInput = document.getElementById('nav-input')
const navReview = document.getElementById('nav-review')

// Add main file input
const mainInput = document.getElementById('main-upload')

let srbankData = null
let duettData = null
let mainData = null
let currentSortMode = 'allByDate' // 'newOnTop' or 'allByDate'

// Save stats to not calculate multiple times
let total_length = 0
