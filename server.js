const express = require('express');
const session = require('express-session');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');
const { Dropbox } = require('dropbox');
const fetch = require('isomorphic-fetch');
require('dotenv').config(); // Charger les variables d'environnement

const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({
    secret: 'secret-key', // changez ceci par une clé secrète sécurisée
    resave: false,
    saveUninitialized: true
}));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'html.html'));
});

app.get('/page2', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'page2.html'));
});

app.get('/submit', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'submit.html'));
});

app.post('/submit-page1', (req, res) => {
    req.session.page1 = req.body;
    res.redirect('/page2');
});

app.post('/submit-page2', async(req, res) => {
    const page1Data = req.session.page1 || {};
    const page2Data = req.body;

    const formData = {...page1Data, ...page2Data };

    const filePath = path.join(__dirname, 'public', 'reponses.xlsx');
    let workbook;
    let worksheet;
    let existingData = [];

    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
        existingData = xlsx.utils.sheet_to_json(worksheet);
    } else {
        workbook = xlsx.utils.book_new();
    }

    existingData.push({
        "Qu'est-ce que signifie la « DATA » ?": formData.question1,
        "Qu'est-ce que le programme Connect de Marjane Group ?": formData.question2,
        "Quelles sont les utilisations de la data mentionnées par le Président ?": formData.question3,
        "Quel est le principal bénéfice mentionné dans la vidéo de Fatim Sefrioui et Hakim Mataich ?": formData.question8,
        "La transformation digitale est-elle un projet à court terme ?": formData.question9,
        "Quelle est la principale raison de la transformation digitale selon Zakaria Essalhi ?": formData.question10,
    });

    worksheet = xlsx.utils.json_to_sheet(existingData);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Réponses');
    xlsx.writeFile(workbook, filePath);

    // Ajouter le code pour uploader vers Dropbox
    const dbx = new Dropbox({ accessToken: process.env.DROPBOX_ACCESS_TOKEN, fetch: fetch });

    try {
        const fileContent = fs.readFileSync(filePath);
        await dbx.filesUpload({ path: '/reponses.xlsx', contents: fileContent, mode: 'overwrite' });
        console.log('Fichier uploadé avec succès sur Dropbox');
    } catch (error) {
        console.error('Erreur lors de l\'upload sur Dropbox :', error);
    }

    res.redirect('/submit');
});

app.listen(3000, () => {
    console.log('Serveur démarré sur http://localhost:3000');
});