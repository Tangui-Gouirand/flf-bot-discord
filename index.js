const { Client, GatewayIntentBits, ButtonBuilder, ActionRowBuilder, ButtonStyle, EmbedBuilder } = require('discord.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent, GatewayIntentBits.GuildMembers] });
const excelFilePath = path.join(__dirname, 'services.xlsx');
const userMessageIds = {}; // Pour stocker les messages avec les boutons

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

    if (message.content.startsWith('!service')) {
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

        // Envoyez le message avec les boutons et stockez l'identifiant du message
        const sentMessage = await message.channel.send({
            content: 'Cliquez sur un bouton pour débuter ou terminer le service.',
            components: [row]
        });

        userMessageIds[sentMessage.id] = { displayName, channelId: message.channel.id }; // Stocke le nom d'affichage et l'ID du canal

        // Supprimer le message d'origine
        await message.delete();
    } else if (message.content.startsWith('!temps')) {
        const args = message.content.split(' ');
        const mentionedUser = message.mentions.users.first();

        if (!mentionedUser) {
            await message.reply("Veuillez mentionner un utilisateur.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const today = new Date().toISOString().slice(0, 10); // Date d'aujourd'hui au format YYYY-MM-DD

        const totalTime = await getTotalTimeWorked(displayName, today);

        const embed = new EmbedBuilder()
            .setTitle(`Temps travaillé aujourd'hui pour ${displayName}`)
            .setDescription(`${displayName} a travaillé un total de ${totalTime} aujourd'hui.`)
            .setColor(0x00FF00); // Vert

        await message.reply({ embeds: [embed] });

        // Supprimer le message d'origine
        await message.delete();
    } else if (message.content.startsWith('!total')) {
        const args = message.content.split(' ');
        const numberOfDays = args[1];
        const mentionedUser = message.mentions.users.first();

        if (!mentionedUser || !numberOfDays || isNaN(numberOfDays)) {
            await message.reply("Veuillez mentionner un utilisateur et spécifier le nombre de jours.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const endDate = new Date();
        const startDate = new Date();
        startDate.setDate(endDate.getDate() - parseInt(numberOfDays));

        const totalTime = await getTotalTimeWorkedInRange(displayName, startDate.toISOString().slice(0, 10), endDate.toISOString().slice(0, 10));

        const embed = new EmbedBuilder()
            .setTitle(`Temps travaillé pour ${displayName} sur les ${numberOfDays} derniers jours`)
            .setDescription(`${displayName} a travaillé un total de ${totalTime}.`)
            .setColor(0x00FF00); // Vert

        await message.reply({ embeds: [embed] });

        // Supprimer le message d'origine
        await message.delete();
    }
});

client.on('interactionCreate', async interaction => {
    if (interaction.isButton()) {
        const userId = interaction.user.id;
        const member = interaction.guild.members.cache.get(userId);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const timestamp = new Date().toLocaleString(); // Date et heure actuelle

        if (interaction.customId === 'startService') {
            await setUserStatus(displayName, 'en service', timestamp);
            const embed = new EmbedBuilder()
                .setColor(0x00FF00) // Vert
                .setDescription(`${displayName} a commencé son service à ${timestamp}.`);

            // Trouver et mettre à jour le message contenant les boutons
            const messageId = Object.keys(userMessageIds).find(id => userMessageIds[id].displayName === displayName);
            if (messageId) {
                const channel = client.channels.cache.get(userMessageIds[messageId].channelId);
                if (channel) {
                    const message = await channel.messages.fetch(messageId);
                    await message.edit({ embeds: [embed], components: [] }); // Met à jour le message avec l'embed et sans boutons
                    delete userMessageIds[messageId]; // Nettoyer l'objet après mise à jour
                }
            }
        } else if (interaction.customId === 'endService') {
            await setUserStatus(displayName, 'hors service', timestamp);
            const embed = new EmbedBuilder()
                .setColor(0xFF0000) // Rouge
                .setDescription(`${displayName} a terminé son service à ${timestamp}.`);

            // Trouver et mettre à jour le message contenant les boutons
            const messageId = Object.keys(userMessageIds).find(id => userMessageIds[id].displayName === displayName);
            if (messageId) {
                const channel = client.channels.cache.get(userMessageIds[messageId].channelId);
                if (channel) {
                    const message = await channel.messages.fetch(messageId);
                    await message.edit({ embeds: [embed], components: [] }); // Met à jour le message avec l'embed et sans boutons
                    delete userMessageIds[messageId]; // Nettoyer l'objet après mise à jour
                }
            }
        }
    } else if (interaction.isStringSelectMenu()) {
        const { customId, values } = interaction;
        const [selectedUser, selectedDays] = values;

        if (customId === 'userSelect') {
            const member = interaction.guild.members.cache.get(selectedUser);
            const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';

            const totalTimeToday = await getTotalTimeWorked(displayName, new Date().toISOString().slice(0, 10));
            const embed = new EmbedBuilder()
                .setTitle(`Temps travaillé aujourd'hui pour ${displayName}`)
                .setDescription(`${displayName} a travaillé un total de ${totalTimeToday} aujourd'hui.`)
                .setColor(0x00FF00); // Vert

            await interaction.update({ embeds: [embed], components: [] });

        } else if (customId === 'daysSelect') {
            const numberOfDays = parseInt(selectedDays);
            const endDate = new Date();
            const startDate = new Date();
            startDate.setDate(endDate.getDate() - numberOfDays);

            const member = interaction.guild.members.cache.get(selectedUser);
            const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';

            const totalTime = await getTotalTimeWorkedInRange(displayName, startDate.toISOString().slice(0, 10), endDate.toISOString().slice(0, 10));
            const embed = new EmbedBuilder()
                .setTitle(`Temps travaillé pour ${displayName} sur les ${numberOfDays} derniers jours`)
                .setDescription(`${displayName} a travaillé un total de ${totalTime}.`)
                .setColor(0x00FF00); // Vert

            await interaction.update({ embeds: [embed], components: [] });
        }
    }
});

async function getUserStatus(displayName) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return 'hors service';

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
            const [timestamp, status] = [row.getCell(1).value, row.getCell(2).value];
            if (timestamp.startsWith(date)) {
                if (status === 'en service') {
                    lastStartTime = new Date(timestamp);
                } else if (status === 'hors service' && lastStartTime) {
                    const endTime = new Date(timestamp);
                    const duration = (endTime - lastStartTime) / 60000; // Convertir la durée en minutes
                    totalMinutes += duration;
                    lastStartTime = null;
                }
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
            const [timestamp, status] = [row.getCell(1).value, row.getCell(2).value];
            if (new Date(timestamp) >= new Date(startDate) && new Date(timestamp) <= new Date(endDate)) {
                if (status === 'en service') {
                    lastStartTime = new Date(timestamp);
                } else if (status === 'hors service' && lastStartTime) {
                    const endTime = new Date(timestamp);
                    const duration = (endTime - lastStartTime) / 60000; // Convertir la durée en minutes
                    totalMinutes += duration;
                    lastStartTime = null;
                }
            }
        });

        const hours = Math.floor(totalMinutes / 60);
        const minutes = Math.round(totalMinutes % 60);
        return `${hours} heures ${minutes} minutes`;
    } catch (error) {
        console.error('Erreur lors du calcul du temps travaillé:', error);
        return 'Erreur';
    }
}

function sanitizeSheetName(name) {
    return name.replace(/[\\/?*[\]:]/g, '_').substring(0, 31);
}

client.login(process.env.TOKEN);



