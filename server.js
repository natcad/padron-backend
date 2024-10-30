const express = require('express');
const nodemailer = require('nodemailer');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
require('dotenv').config({ path: '.env' }); // Cargar var.env

const app = express();

app.use(cors({
    origin: 'https://padron-coronel-suarez.onrender.com', // Cambia esto por la URL de tu frontend
    methods: ['GET', 'POST'],
    allowedHeaders: ['Content-Type'],
}));
app.use(express.json());

// Configura el transportador de nodemailer
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.GMAIL_USER,
        pass: process.env.GMAIL_PASS,
    },
});

// Array para almacenar las consultas y el tiempo de envio
let consultas = [];
let lastSentTime = null; 

// Función para enviar correos electrónicos
const sendEmail = async (from, subject, body, attachmentPath) => {
    const mailOptions = {
        from,
        to: process.env.GMAIL_RECIPIENT,
        subject,
        text: body,
        attachments: attachmentPath ? [{ path: attachmentPath }] : [], // Adjunta el archivo si se proporciona
    };

    try {
        const info = await transporter.sendMail(mailOptions);
        console.log('Correo enviado: ', info.response);
    } catch (error) {
        console.error('Error al enviar el correo: ', error);
    }
};

// Crear o cargar un archivo Excel existente
const createOrLoadExcel = () => {
    const filePath = path.join(__dirname, 'consultas.xlsx');
    let data = [];

    // Verifica si el archivo ya existe
    if (fs.existsSync(filePath)) {
        const workbook = XLSX.readFile(filePath);
        const worksheet = workbook.Sheets['Consultas'];
        data = XLSX.utils.sheet_to_json(worksheet);
    }

    consultas = data; // Cargar datos existentes en el array consultas
    return filePath;
};

// Función para verificar si la consulta ya existe
const consultaExists = (nombre, apellido, dni) => {
    return consultas.some((consulta) => {
        return consulta.nombre === nombre && consulta.apellido === apellido && consulta.dni === dni;
    });
};

// Función para actualizar el archivo Excel con las consultas acumuladas
const updateExcelWithConsultas = () => {
    const filePath = path.join(__dirname, 'consultas.xlsx');
    const wb = XLSX.utils.book_new();

    // Si hay consultas, añade una nueva hoja
    if (consultas.length > 0) {
        const ws = XLSX.utils.json_to_sheet(consultas);
        XLSX.utils.book_append_sheet(wb, ws, 'Consultas');
    }

    XLSX.writeFile(wb, filePath);
    return filePath;
};
const checkAndSendEmail = () => {
    const currentTime = new Date();
    
    // Verificar si ha pasado una semana desde el último envío
    const oneWeek = 7 * 24 * 60 * 60 * 1000; // 1 semana en milisegundos
    const hasOneWeekPassed = lastSentTime === null || (currentTime - lastSentTime >= oneWeek);
    
    // Verificar si se han hecho al menos 10 consultas
    if (consultas.length >= 50 && (hasOneWeekPassed || lastSentTime === null)) {
        // Actualizar el archivo Excel con las consultas
        const updatedExcelFilePath = updateExcelWithConsultas();

        // Enviar correo con el archivo actualizado
        sendEmail(process.env.GMAIL_USER, 'Consulta de Afiliado', `Este es el archivo Excel con las consultas actualizadas de las fechas ${lastSentTime} - ${currentTime} .`, updatedExcelFilePath);
        
        // Actualizar el tiempo del último envío
        lastSentTime = currentTime;
        
        // Restablecer las consultas a un array vacío
        consultas = [];
    }
};

app.post('/send-email',  cors(), (req, res) => {
    const { nombre, apellido, dni, afiliado } = req.body;

    // Verificar si la consulta ya existe
    if (consultaExists(nombre, apellido, dni)) {
        return res.status(400).send('La consulta ya existe.');
    }

    // Agregar nueva consulta
    consultas.push({ nombre, apellido, dni, afiliado });

    // Verificar si se debe enviar el correo
    checkAndSendEmail();

    res.status(200).send('Consulta recibida.');
});

// Inicializar el Excel y enviar un correo inicial si es necesario
createOrLoadExcel();

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});