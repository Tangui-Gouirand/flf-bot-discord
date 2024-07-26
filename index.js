const { Client, GatewayIntentBits, ButtonBuilder, ActionRowBuilder, ButtonStyle, EmbedBuilder, Events } = require('discord.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent, GatewayIntentBits.GuildMembers] });

const excelFilePath = path.join(__dirname, 'services.xlsx');
const userMessageIds = {};

// Fonction pour créer un fichier Excel vide s'il n'existe pas
async function createExcelFileIfNotExists() {
    try {
        if (!fs.existsSync(excelFilePath)) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.writeFile(excelFilePath);
            console.log('Fichier Excel créé.');
        }
    } catch (error) {
        console.error('Erreur lors de la création du fichier Excel:', error);
    }
}

// Assurez-vous que le fichier Excel existe lorsque le bot démarre
client.once('ready', async () => {
    await createExcelFileIfNotExists();
    console.log('Bot prêt!');
});

client.on('messageCreate', async message => {
    if (message.author.bot) return; // Ignore les messages des bots

    if (message.content.startsWith('!help')) {
        const helpMessage = `
**Commandes disponibles :**

**!service**
- Affiche les boutons pour débuter ou terminer le service.

**!temps @USER**
- Affiche le temps travaillé aujourd'hui pour l'utilisateur mentionné.

**!total [N] @USER**
- Affiche le temps total travaillé par l'utilisateur mentionné durant les N derniers jours.

**!help**
- Affiche cette aide.

Utilisez ces commandes pour gérer et consulter les temps de service.
`;

        await message.reply(helpMessage);
    } else if (message.content.startsWith('!service')) {
        const userId = message.author.id;
        const member = message.guild.members.cache.get(userId);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const userStatus = await getUserStatus(displayName);

        const startButton = new ButtonBuilder()
            .setCustomId('startService')
            .setLabel('Début de Service')
            .setStyle(ButtonStyle.Primary)
            .setDisabled(userStatus === 'en service');

        const endButton = new ButtonBuilder()
            .setCustomId('endService')
            .setLabel('Fin de Service')
            .setStyle(ButtonStyle.Danger)
            .setDisabled(userStatus !== 'en service');

        const row = new ActionRowBuilder().addComponents(startButton, endButton);

        // Envoyez le message avec les boutons et stockez l'identifiant de l'utilisateur
        const sentMessage = await message.channel.send({
            content: 'Cliquez sur un bouton pour débuter ou terminer le service.',
            components: [row]
        });

        userMessageIds[sentMessage.id] = displayName;
    } else if (message.content.startsWith('!temps')) {
        // Récupère l'utilisateur mentionné
        const args = message.content.split(' ');
        const mentionedUser = message.mentions.users.first();

        if (!mentionedUser) {
            await message.reply("Veuillez mentionner un utilisateur.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const today = new Date().toISOString().slice(0, 10); // Date d'aujourd'hui au format YYYY-MM-DD

        // Récupère le temps travaillé aujourd'hui pour l'utilisateur mentionné
        const totalTime = await getTotalTimeWorked(displayName, today);

        const embed = new EmbedBuilder()
            .setTitle(`Temps travaillé aujourd'hui pour ${displayName}`)
            .setDescription(`${displayName} a travaillé un total de ${totalTime} aujourd'hui.`)
            .setColor(0x00FF00); // Vert

        await message.reply({ embeds: [embed] });
    } else if (message.content.startsWith('!total')) {
        // Traitement de la commande !total
        const args = message.content.split(' ');
        const days = parseInt(args[1], 10);
        const mentionedUser = message.mentions.users.first();

        if (!mentionedUser || isNaN(days)) {
            await message.reply("Veuillez mentionner un utilisateur et spécifier le nombre de jours.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';

        const endDate = new Date();
        const startDate = new Date();
        startDate.setDate(endDate.getDate() - days);
        const startDateStr = startDate.toISOString().slice(0, 10); // Format YYYY-MM-DD
        const endDateStr = endDate.toISOString().slice(0, 10); // Format YYYY-MM-DD

        // Récupère le temps total travaillé pour l'utilisateur mentionné sur la période donnée
        const totalTime = await getTotalTimeWorkedInRange(displayName, startDateStr, endDateStr);

        const embed = new EmbedBuilder()
            .setTitle(`Temps travaillé pour ${displayName} du ${startDateStr} au ${endDateStr}`)
            .setDescription(`${displayName} a travaillé un total de ${totalTime} pendant cette période.`)
            .setColor(0x00FF00); // Vert

        await message.reply({ embeds: [embed] });
    }
});

client.on(Events.InteractionCreate, async interaction => {
    if (!interaction.isButton()) return;

    const userId = interaction.user.id;
    const member = interaction.guild.members.cache.get(userId);
    const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';

    // Vérifiez si l'utilisateur qui a initié la commande est le même que celui qui clique sur les boutons
    if (userMessageIds[interaction.message.id] !== displayName) {
        await interaction.reply({ content: "Vous ne pouvez pas utiliser ce bouton.", ephemeral: true });
        return;
    }

    let embed;
    const timestamp = new Date().toLocaleString(); // Date et heure actuelle
    if (interaction.customId === 'startService') {
        await setUserStatus(displayName, 'en service', timestamp);
        embed = new EmbedBuilder()
            .setColor(0x00FF00) // Vert
            .setDescription(`${displayName} a commencé son service à ${timestamp}.`);
    } else if (interaction.customId === 'endService') {
        const userStatus = await getUserStatus(displayName);
        if (userStatus === 'en service') {
            await setUserStatus(displayName, 'hors service', timestamp);
            embed = new EmbedBuilder()
                .setColor(0xFF0000) // Rouge
                .setDescription(`${displayName} a terminé son service à ${timestamp}.`);
        } else {
            embed = new EmbedBuilder()
                .setColor(0xFF0000) // Rouge
                .setDescription(`${displayName} n'est pas en service.`);
        }
    }

    await interaction.reply({ embeds: [embed] });

    // Mettre à jour les boutons après l'interaction
    const userStatus = await getUserStatus(displayName);
    const startButton = new ButtonBuilder()
        .setCustomId('startService')
        .setLabel('Début de Service')
        .setStyle(ButtonStyle.Primary)
        .setDisabled(userStatus === 'en service');

    const endButton = new ButtonBuilder()
        .setCustomId('endService')
        .setLabel('Fin de Service')
        .setStyle(ButtonStyle.Danger)
        .setDisabled(userStatus !== 'en service');

    const row = new ActionRowBuilder().addComponents(startButton, endButton);

    await interaction.message.edit({
        components: [row]
    });
});

async function getUserStatus(displayName) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return 'hors service';

        // Trouver la dernière ligne
        const lastRow = sheet.lastRow;
        if (lastRow && lastRow.getCell(2).value === 'en service') {
            return 'en service';
        }
        return 'hors service';
    } catch (error) {
        console.error('Erreur lors de la lecture du statut utilisateur:', error);
        return 'hors service';
    }
}

async function setUserStatus(displayName, status, timestamp) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        let sheet = workbook.getWorksheet(displayName);
        if (!sheet) {
            sheet = workbook.addWorksheet(displayName);
            sheet.addRow(['Timestamp', 'Status']); // Ajouter les en-têtes de colonne
        }

        // Ajouter l'enregistrement du service
        sheet.addRow([timestamp, status]);

        await workbook.xlsx.writeFile(excelFilePath);
    } catch (error) {
        console.error('Erreur lors de l\'enregistrement du statut utilisateur:', error);
    }
}

async function getTotalTimeWorked(displayName, date) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return '0 heures 0 minutes';

        let totalMinutes = 0;
        let lastStartTime = null;

        sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const timestamp = row.getCell(1).value;
            const status = row.getCell(2).value;

            if (timestamp && status) {
                const rowDate = parseDate(timestamp);

                // Vérifier si la date est valide
                if (isNaN(rowDate.getTime())) {
                    console.warn(`Valeur de date invalide trouvée: ${timestamp}`);
                    return;
                }

                const rowDateStr = rowDate.toISOString().slice(0, 10);

                if (rowDateStr === date) {
                    if (status === 'en service') {
                        lastStartTime = rowDate;
                    } else if (status === 'hors service' && lastStartTime) {
                        const endTime = rowDate;
                        const duration = (endTime - lastStartTime) / 60000; // Convertir la durée en minutes
                        totalMinutes += duration;
                        lastStartTime = null; // Réinitialiser la dernière heure de début
                    }
                }
            } else {
                console.warn(`Ligne avec des données manquantes: ${row.values}`);
            }
        });

        const hours = Math.floor(totalMinutes / 60);
        const minutes = Math.round(totalMinutes % 60);
        return `${hours} heures ${minutes} minutes`;
    } catch (error) {
        console.error('Erreur lors de la lecture du temps travaillé:', error);
        return 'Erreur';
    }
}

async function getTotalTimeWorkedInRange(displayName, startDate, endDate) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return '0 heures 0 minutes';

        let totalMinutes = 0;
        let lastStartTime = null;

        sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const timestamp = row.getCell(1).value;
            const status = row.getCell(2).value;

            if (timestamp && status) {
                const rowDate = parseDate(timestamp);

                // Vérifier si la date est valide
                if (isNaN(rowDate.getTime())) {
                    console.warn(`Valeur de date invalide trouvée: ${timestamp}`);
                    return;
                }

                const rowDateStr = rowDate.toISOString().slice(0, 10);

                if (rowDateStr >= startDate && rowDateStr <= endDate) {
                    if (status === 'en service') {
                        lastStartTime = rowDate;
                    } else if (status === 'hors service' && lastStartTime) {
                        const endTime = rowDate;
                        const duration = (endTime - lastStartTime) / 60000; // Convertir la durée en minutes
                        totalMinutes += duration;
                        lastStartTime = null; // Réinitialiser la dernière heure de début
                    }
                }
            } else {
                console.warn(`Ligne avec des données manquantes: ${row.values}`);
            }
        });

        const hours = Math.floor(totalMinutes / 60);
        const minutes = Math.round(totalMinutes % 60);
        return `${hours} heures ${minutes} minutes`;
    } catch (error) {
        console.error('Erreur lors de la lecture du temps travaillé:', error);
        return 'Erreur';
    }
}

// Fonction pour convertir une date dans le format DD/MM/YYYY HH:MM:SS en format Date JavaScript
function parseDate(dateString) {
    if (typeof dateString !== 'string') {
        console.warn(`Date reçue n'est pas une chaîne de caractères: ${dateString}`);
        return new Date(NaN); // Retourne une date invalide
    }
    
    const parts = dateString.split(' ');
    if (parts.length !== 2) {
        console.warn(`Format de date invalide: ${dateString}`);
        return new Date(NaN); // Retourne une date invalide
    }

    const dateParts = parts[0].split('/');
    const timeParts = parts[1].split(':');
    
    if (dateParts.length !== 3 || timeParts.length !== 3) {
        console.warn(`Format de date invalide: ${dateString}`);
        return new Date(NaN); // Retourne une date invalide
    }

    const [day, month, year] = dateParts;
    const [hour, minute, second] = timeParts;
    
    return new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}`);
}

// Fonction pour nettoyer le nom d'affichage afin qu'il soit valide comme nom de feuille Excel
function sanitizeSheetName(name) {
    return name.replace(/[\\/?*[\]:]/g, '_').substring(0, 31);
}


client.login(process.env.TOKEN);
