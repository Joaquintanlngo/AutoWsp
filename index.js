const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const xlsx = require('xlsx');

// Leer Excel
const workbook = xlsx.readFile('numeros.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

// Crear cliente
const client = new Client({
    authStrategy: new LocalAuth()
});

const fs = require('fs');

let results = [];
const sentNumbers = new Set();

// 📸 CARGAR IMAGEN
// const media = MessageMedia.fromFilePath('./images/fondo.jpg');

// Si el archivo existe, lo leemos
if (fs.existsSync('resultados_envio.xlsx')) {
    const workbookExistente = xlsx.readFile('resultados_envio.xlsx');
    const sheetExistente = workbookExistente.Sheets[workbookExistente.SheetNames[0]];
    results = xlsx.utils.sheet_to_json(sheetExistente);
}

function saveResult(newData) {
    const index = results.findIndex(
        r => r.Nombre.toLowerCase().trim() ===
            newData.Nombre.toLowerCase().trim() &&
            r.Colegio.toLowerCase().trim() ===
            newData.Colegio.toLowerCase().trim()
    );

    if (index !== -1) {
        results[index].Telefono = newData.Telefono;
        results[index].Estado = newData.Estado;
        results[index].Colegio = newData.Colegio;
        results[index].Fecha = new Date().toLocaleString();
    } else {
        results.push({
            ...newData,
            Fecha: new Date().toLocaleString()
        });
    }
}

function formatPhone(phone) {
    const clean = phone.replace(/^549/, '');
    const area = clean.slice(0, 3);
    const part1 = clean.slice(3, 6);
    const part2 = clean.slice(6);

    return `${area} ${part1}-${part2}`;
}

// QR
client.on('qr', qr => {
    console.log('Escaneá este QR con WhatsApp:');
    qrcode.generate(qr, { small: true });
});

// MENSAJES
const messages = [
`Hola, ¿cómo están?
Mi nombre es Joaquín Tanlongo, me contacto desde Junior Achievement. Soy coordinador de los programas "Yo Puedo Programar" e "Inteligencia Artificial" y quería consultar si están interesados en participar de nuestras propuestas educativas.`,

`Buen día
Me presento, soy Joaquín Tanlongo y trabajo en Junior Achievement como coordinador de los programas YPP e Inteligencia Artificial. Quisiera saber si les interesa que sus estudiantes participen de nuestros programas educativos.`,

`Buenas, ¿cómo están?
Soy Joaquín Tanlongo, coordinador de los programas de programación e IA en Junior Achievement. Me contacto para consultar si les interesaría sumarse a nuestras capacitaciones.`,

`Qué tal, ¿cómo están?
Mi nombre es Joaquín Tanlongo, trabajo en Junior Achievement y estoy a cargo de los programas Yo Puedo Programar e Inteligencia Artificial. Quería consultar si están interesados en participar de nuestras iniciativas educativas.`,

`Hola profe, ¿cómo está?
Me presento, soy Joaquín Tanlongo, coordinador de los programas YPP e IA en Junior Achievement. Me contacto para saber si les interesaría participar con sus estudiantes en nuestras oportunidades educativas.`,

`Buen día, ¿cómo están?
Soy Joaquín Tanlongo, trabajo en Junior Achievement como coordinador de los programas Yo Puedo Programar e Inteligencia Artificial. Quería consultar si están interesados en participar de nuestros programas de formación en habilidades digitales.`,

`Buenas
Mi nombre es Joaquín Tanlongo, me contacto desde Junior Achievement. Estoy a cargo de los programas YPP e Inteligencia Artificial y quería saber si les interesa sumarse a nuestras experiencias educativas.`,

`Hola, ¿cómo va?
Soy Joaquín Tanlongo, coordinador de los programas Yo Puedo Programar e Inteligencia Artificial en Junior Achievement. Me contacto para consultar si están interesados en participar de nuestros programas gratuitos.`
];

// 📩 SEGUNDO MENSAJE (IMAGEN + TEXTO)
// const secondMessage = `📢 Te comparto más información sobre nuestros programas:

// 🤖 Inteligencia Artificial  
// 💻 Yo Puedo Programar  

// ✅ Gratuitos  
// ✅ Modalidad virtual  
// ✅ Para estudiantes desde 15 años  

// 🚀 Ya participaron miles de estudiantes en todo el país  

// Si te interesa, te puedo pasar el link de inscripción 😊`;

function getRandomMessage() {
    return messages[Math.floor(Math.random() * messages.length)];
}

function getRandomDelay() {
    return Math.floor(Math.random() * (45000 - 15000) + 15000);
}

let count = 0;
let lastSchool = '';
let alreadyStarted = false;

// READY
client.on('ready', async () => {
    if (alreadyStarted) {
        console.log("⚠️ Ready duplicado detectado, ignorando...");
        return;
    }

    alreadyStarted = true;

    console.log('✅ WhatsApp listo!');

    for (let i = 0; i < data.length; i++) {
        const name = data[i].Nombre;
        let school = data[i].Colegio;

        if (!school) {
            school = lastSchool;
        } else {
            lastSchool = school;
        }

        let rawPhone = String(data[i].Telefono).replace(/\D/g, '');

        if (rawPhone.startsWith('0')) rawPhone = rawPhone.substring(1);
        if (rawPhone.startsWith('5490')) rawPhone = rawPhone.substring(4);
        if (rawPhone.startsWith('15')) rawPhone = rawPhone.substring(2);

        let phone = rawPhone;
        if (!rawPhone.startsWith('549')) phone = '549' + rawPhone;

        if (phone.length < 13) {
            console.log(`❌ ${name} número inválido ${school}`);
            saveResult({
                Colegio: school,
                Nombre: name,
                Telefono: formatPhone(phone),
                Estado: 'NUMERO INVALIDO'
            });
            continue;
        }

        const number = phone + '@c.us';

        if (sentNumbers.has(number)) {
            console.log(`⏭️ Ya enviado: ${name}`);
            continue;
        }

        sentNumbers.add(number);

        try {
            const isRegistered = await client.isRegisteredUser(number);

            if (!isRegistered) {
                console.log(`❌ ${name} no tiene WhatsApp`);
                saveResult({
                    Colegio: school,
                    Nombre: name,
                    Telefono: formatPhone(phone),
                    Estado: 'SIN WHATSAPP'
                });
                continue;
            }

            const message = getRandomMessage();
            const chat = await client.getChatById(number);

            // 🧠 PRIMER MENSAJE
            await chat.sendStateTyping();
            await new Promise(r => setTimeout(r, 2000 + Math.random() * 2000));

            await client.sendMessage(number, message);

            // 🧠 SEGUNDO MENSAJE (IMAGEN)
            // await new Promise(r => setTimeout(r, 4000 + Math.random() * 4000));

            // await chat.sendStateTyping();
            // await new Promise(r => setTimeout(r, 2000 + Math.random() * 2000));

            // await client.sendMessage(number, media, {
            //     caption: secondMessage
            // });

            console.log(`✅ Mensaje enviado a ${name} - ${school}`);

            saveResult({
                Colegio: school,
                Nombre: name,
                Telefono: formatPhone(phone),
                Estado: 'ENVIADO'
            });

            count++;

            const delay = getRandomDelay();
            console.log(`⏳ Esperando ${Math.round(delay / 1000)}s...`);
            await new Promise(r => setTimeout(r, delay));

        } catch (error) {
            console.log(`❌ Error con ${name}`);
            saveResult({
                Colegio: school,
                Nombre: name,
                Telefono: formatPhone(phone),
                Estado: 'ERROR'
            });
        }
    }

    console.log(`🎉 Todos los mensajes enviados a ${count} contactos!!`);

    const newSheet = xlsx.utils.json_to_sheet(results);

    newSheet['!cols'] = [
        { wch: 50 },
        { wch: 25 },
        { wch: 20 },
        { wch: 20 },
        { wch: 25 }
    ];

    const newBook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newBook, newSheet, 'Resultados');
    xlsx.writeFile(newBook, 'resultados_envio.xlsx');

    await client.destroy();
    process.exit();
});

client.initialize();