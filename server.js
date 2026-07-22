const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Se il file prono.xlsx non esiste, usa dati di esempio
let matchesData = [];

function loadData() {
    try {
        const filePath = path.join(__dirname, 'prono.xlsx');
        if (!fs.existsSync(filePath)) {
            console.log('⚠️ File prono.xlsx non trovato, uso dati di esempio');
            return generateSampleData();
        }
        
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets['Partite'];
        const data = XLSX.utils.sheet_to_json(sheet);
        return data;
    } catch (error) {
        console.error('Errore caricamento dati:', error);
        return generateSampleData();
    }
}

function generateSampleData() {
    const today = new Date().toISOString().split('T')[0];
    const tomorrow = new Date(Date.now() + 86400000).toISOString().split('T')[0];
    const dayAfter = new Date(Date.now() + 172800000).toISOString().split('T')[0];
    
    return [
        { Campionato: 'Eliteserien', Giornata: '1', Data: today, Ora: '19:00', 'Squadra Casa': 'Bodø/Glimt', 'Squadra Ospite': 'Rosenborg' },
        { Campionato: 'Eliteserien', Giornata: '1', Data: today, Ora: '19:00', 'Squadra Casa': 'Viking', 'Squadra Ospite': 'Molde' },
        { Campionato: 'Eliteserien', Giornata: '1', Data: tomorrow, Ora: '18:00', 'Squadra Casa': 'Brann', 'Squadra Ospite': 'Vålerenga' },
        { Campionato: 'Allsvenskan', Giornata: '13', Data: today, Ora: '16:30', 'Squadra Casa': 'Halmstad', 'Squadra Ospite': 'BK Häcken' },
        { Campionato: 'Allsvenskan', Giornata: '13', Data: today, Ora: '16:30', 'Squadra Casa': 'Elfsborg', 'Squadra Ospite': 'Sirius' },
        { Campionato: 'Allsvenskan', Giornata: '13', Data: today, Ora: '16:30', 'Squadra Casa': 'Hammarby', 'Squadra Ospite': 'Degerfors' },
    ];
}

function getDates() {
    const today = new Date();
    const dates = [];
    for (let i = 0; i < 3; i++) {
        const d = new Date(today);
        d.setDate(d.getDate() + i);
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        dates.push(`${year}-${month}-${day}`);
    }
    return dates;
}

app.get('/api/matches', (req, res) => {
    try {
        const data = loadData();
        const targetDates = getDates();

        const matches = data
            .filter(row => {
                const dateStr = String(row.Data || row['Data'] || '').trim();
                if (!dateStr) return false;
                return targetDates.some(target => dateStr.includes(target));
            })
            .map(row => ({
                campionato: String(row.Campionato || row['Campionato'] || ''),
                giornata: String(row.Giornata || row['Turno'] || row['Giornata'] || ''),
                data: String(row.Data || row['Data'] || ''),
                ora: String(row.Ora || row['Ora'] || ''),
                casa: String(row['Squadra Casa'] || row['Squadra'] || row['Home'] || ''),
                ospite: String(row['Squadra Ospite'] || row['Ospite'] || row['Away'] || ''),
                gol_casa: parseInt(row['Gol Casa'] || row['GC'] || 0),
                gol_ospite: parseInt(row['Gol Ospite'] || row['GO'] || 0)
            }))
            .filter(m => m.campionato && m.casa && m.ospite);

        matches.sort((a, b) => (a.data + a.ora).localeCompare(b.data + b.ora));

        res.json({
            success: true,
            matches: matches,
            total: matches.length,
            dates: targetDates
        });

    } catch (error) {
        console.error('Errore:', error);
        res.status(500).json({ success: false, error: error.message });
    }
});

app.get('/api/standings', (req, res) => {
    try {
        const filePath = path.join(__dirname, 'prono.xlsx');
        if (!fs.existsSync(filePath)) {
            return res.json({ success: true, standings: [] });
        }
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets['Classifica'];
        const data = XLSX.utils.sheet_to_json(sheet);
        res.json({ success: true, standings: data });
    } catch (error) {
        res.json({ success: true, standings: [] });
    }
});

// Serve index.html per tutte le altre route
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`✅ Server avviato sulla porta ${PORT}`);
    console.log(`📁 Cartella: ${__dirname}`);
    console.log(`🌐 Accessibile da tutti`);
});