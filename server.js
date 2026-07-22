const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname)));

// Funzione per ottenere le date
function getDates() {
    const today = new Date();
    const dates = [];
    // Formato AAAA-MM-GG (quello che usa il server)
    const formatISO = (d) => {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };
    // Formato GG/MM/AAAA (quello che hai nel tuo Excel)
    const formatEU = (d) => {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${day}/${month}/${year}`;
    };
    
    for (let i = 0; i < 3; i++) {
        const d = new Date(today);
        d.setDate(d.getDate() + i);
        dates.push(formatISO(d));
        dates.push(formatEU(d));  // Aggiungi anche il formato EU
    }
    return dates;
}

// Funzione per leggere il file Excel
function readExcelFile(filePath, sheetName) {
    try {
        if (!fs.existsSync(filePath)) {
            return null;
        }
        const workbook = XLSX.readFile(filePath);
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return null;
        return XLSX.utils.sheet_to_json(sheet);
    } catch (error) {
        console.error(`Errore lettura ${sheetName}:`, error);
        return null;
    }
}

// ENDPOINT: PARTITE
app.get('/api/matches', (req, res) => {
    try {
        const filePath = path.join(__dirname, 'prono.xlsx');
        
        if (!fs.existsSync(filePath)) {
            return res.json({ 
                success: true, 
                matches: [], 
                warning: 'File prono.xlsx non trovato' 
            });
        }

        const data = readExcelFile(filePath, 'Partite');
        if (!data) {
            return res.json({ 
                success: true, 
                matches: [], 
                warning: 'Foglio "Partite" non trovato' 
            });
        }

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

        console.log(`✅ Trovate ${matches.length} partite`);
        
        res.json({
            success: true,
            matches: matches,
            dates: targetDates,
            total: matches.length
        });

    } catch (error) {
        console.error('❌ Errore /api/matches:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ENDPOINT: CLASSIFICA
app.get('/api/standings', (req, res) => {
    try {
        const filePath = path.join(__dirname, 'prono.xlsx');
        
        if (!fs.existsSync(filePath)) {
            return res.json({ success: true, standings: [] });
        }

        const data = readExcelFile(filePath, 'Classifica');
        if (!data) {
            return res.json({ success: true, standings: [] });
        }
        
        res.json({
            success: true,
            standings: data,
            total: data.length
        });
    } catch (error) {
        console.error('❌ Errore /api/standings:', error);
        res.status(500).json({
            success: false,
            error: error.message
        });
    }
});

// ENDPOINT: TEST
app.get('/api/test', (req, res) => {
    const filePath = path.join(__dirname, 'prono.xlsx');
    res.json({
        success: true,
        fileExists: fs.existsSync(filePath),
        filePath: filePath,
        cwd: __dirname
    });
});

// AVVIO SERVER
app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n✅ Server avviato sulla porta ${PORT}`);
    console.log(`📁 Cartella: ${__dirname}`);
    console.log(`📊 prono.xlsx: ${fs.existsSync(path.join(__dirname, 'prono.xlsx')) ? '✅ TROVATO' : '❌ NON TROVATO'}`);
    console.log(`\n🌐 Accessibile da tutti\n`);
});