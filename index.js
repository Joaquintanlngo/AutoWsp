const { Client, LocalAuth } = require('whatsapp-web.js');
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
        // ✔ EXISTE → actualizar solo lo necesario
        results[index].Telefono = newData.Telefono;
        results[index].Estado = newData.Estado;
        results[index].Colegio = newData.Colegio;
        results[index].Fecha = new Date().toLocaleString();
    } else {
        // ✔ NO EXISTE → crear nuevo
        results.push({
            ...newData,
            Fecha: new Date().toLocaleString()
        });
    }
}

function formatPhone(phone) {
    const clean = phone.replace(/^549/, ''); // sacar 549

    // if (clean.length === 10) {
    const area = clean.slice(0, 3);
    const part1 = clean.slice(3, 6);
    const part2 = clean.slice(6);

    return `${area} ${part1}-${part2}`;
    // }

    // return clean; // fallback si no cumple formato
}

// QR
client.on('qr', qr => {
    console.log('Escaneá este QR con WhatsApp:');
    qrcode.generate(qr, { small: true });
});


// Creamos el mensaje

const mensaje1 = "Hola soy joaco, este es mi telefono del laburo"
const mensaje2 = "Estoy probando unos mensajes automaticos";

let count = 0;
let lastSchool = '';
// Cuando está listo
client.on('ready', async () => {
    console.log('✅ WhatsApp listo!');

    for (let i = 0; i < data.length; i++) {
        const name = data[i].Nombre;
        let school = data[i].Colegio;

        // 🧠 Si viene undefined, usamos el anterior
        if (!school) {
            school = lastSchool;
        } else {
            lastSchool = school;
        }
        let rawPhone = String(data[i].Telefono).replace(/\D/g, '');


        // quitar 0 inicial si existe (ej: 0341...)
        if (rawPhone.startsWith('0')) {
            rawPhone = rawPhone.substring(1);
        }
        if (rawPhone.startsWith('5490')) {
            rawPhone = rawPhone.substring(4);
        }

        // quitar 15 si viene como 34115...
        if (rawPhone.startsWith('15')) {
            rawPhone = rawPhone.substring(2);
        }

        // armar formato final Argentina
        let phone = rawPhone;

        if (!rawPhone.startsWith('549')) {
            phone = '549' + rawPhone;
        }

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

            await client.sendMessage(number, mensaje1);

            await new Promise(resolve => setTimeout(resolve, 2000));

            await client.sendMessage(number, mensaje2);

            console.log(`✅ 2 mensajes enviados a ${name} del colegio: ${school}`);
            saveResult({
                Colegio: school,
                Nombre: name,
                Telefono: formatPhone(phone),
                Estado: 'ENVIADO'
            });
            count++;

            await new Promise(resolve => setTimeout(resolve, 3000));

        } catch (error) {
            console.log(`❌ ${name} número inválido o sin WhatsApp`);
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

    // Ajustar ancho de columnas
    newSheet['!cols'] = [
        { wch: 50 }, // Colegio
        { wch: 25 }, // Name
        { wch: 20 }, // Phone
        { wch: 20 }, // Status
        { wch: 25 }  // Date
    ];
    const newBook = xlsx.utils.book_new();

    xlsx.utils.book_append_sheet(newBook, newSheet, 'Resultados');

    xlsx.writeFile(newBook, 'resultados_envio.xlsx');

    // cerrar sesión / proceso
    await client.destroy();
    process.exit();
});

client.initialize();