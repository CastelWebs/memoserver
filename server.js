const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const app = express();

app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');  // Permite todas las solicitudes
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    next();
});

app.use(bodyParser.json());

app.post('/registrar', (req, res) => {
    const { nombre, email, phone , company } = req.body;
    console.log("Usuario Registrado")
    let workbook;
    const filePath = './usuarios.xlsx';
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.aoa_to_sheet([['Nombre', 'Email', 'Telefono', 'Compañia']]); 
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Usuarios');
    }

    const worksheet = workbook.Sheets['Usuarios'];

    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    jsonData.push([nombre, email, phone, company]);

    const newWorksheet = xlsx.utils.aoa_to_sheet(jsonData);
    workbook.Sheets['Usuarios'] = newWorksheet;

    xlsx.writeFile(workbook, filePath);

    res.status(200).json({ message: 'Datos guardados en usuarios.xlsx' });
});

app.post('/actualizarVictoria', (req, res) => {
    const { email } = req.body; // Email del usuario a buscar
    console.log("Actualizando Victoria para el usuario");

    const filePath = './usuarios.xlsx';
    let workbook;

    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        return res.status(404).json({ message: 'El archivo usuarios.xlsx no existe' });
    }

    const worksheet = workbook.Sheets['Usuarios'];
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const header = jsonData[0]; 
    const victoriaIndex = header.indexOf("Victoria");

    if (victoriaIndex === -1) {
        header.push("Victoria");
        jsonData[0] = header;
    }

    const userIndex = jsonData.findIndex(row => row[header.indexOf('Email')] === email);

    if (userIndex === -1) {
        return res.status(404).json({ message: 'Usuario no encontrado' });
    }

    jsonData[userIndex][header.indexOf('Victoria')] = true;

    const newWorksheet = xlsx.utils.aoa_to_sheet(jsonData);
    workbook.Sheets['Usuarios'] = newWorksheet;
    xlsx.writeFile(workbook, filePath);

    res.status(200).json({ message: 'Victoria actualizada a true para el usuario con email: ' + email });
});


app.listen(3000, () => {
    console.log('Servidor corriendo en http://localhost:3000');
});

