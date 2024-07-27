const { Client, GatewayIntentBits, EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } = require('discord.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const moment = require('moment');
require('dotenv').config();

const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent, GatewayIntentBits.GuildMembers] });
const excelFilePath = path.join(__dirname, 'services.xlsx');
const userMessageData = {};

// Fonction pour créer un fichier Excel vide s'il n'existe pas
async function createExcelFileIfNotExists() {
    if (!fs.existsSync(excelFilePath)) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.writeFile(excelFilePath);
    }
}

client.once('ready', async () => {
    await createExcelFileIfNotExists();
    console.log('Bot prêt!');
});

// Fonction pour obtenir le statut de l'utilisateur
async function getUserStatus(displayName) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet || !sheet.lastRow) return 'hors service';
        return sheet.lastRow.getCell(2).value || 'hors service';
    } catch (error) {
        console.error('Erreur lors de la récupération du statut de l\'utilisateur:', error);
        return 'Erreur';
    }
}

// Fonction pour obtenir l'historique des services
async function getServiceHistory(displayName) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return 'Aucun historique trouvé.';
        
        let history = '';
        sheet.eachRow({ includeEmpty: true }, (row) => {
            const [timestamp, status] = [row.getCell(1).value, row.getCell(2).value];
            history += `\n${formatDate(new Date(timestamp))} - ${status}`;
        });
        return history || 'Aucun historique trouvé.';
    } catch (error) {
        console.error('Erreur lors de la récupération de l\'historique des services:', error);
        return 'Erreur';
    }
}

function formatDate(date) {
    return moment(date).format('dddd, MMMM Do YYYY, h:mm:ss a');
}

function createServiceEmbed(displayName, status, timestamp) {
    const color = status === 'en service' ? 0x00FF00 : 0xFF0000;
    const action = status === 'en service' ? 'débuté' : 'terminé';
    return new EmbedBuilder()
        .setColor(color)
        .setDescription(`Service ${action} pour ${displayName} à ${formatDate(timestamp)}.`);
}

function createHelpEmbed() {
    return new EmbedBuilder()
        .setTitle('Commandes du Bot')
        .addFields(
            { name: '!service', value: 'Commence ou termine un service.', inline: true },
            { name: '!temps @USER', value: 'Affiche le temps travaillé aujourd\'hui pour un utilisateur.', inline: true },
            { name: '!total [N] @USER', value: 'Affiche le temps travaillé pour un utilisateur sur les derniers jours.', inline: true },
            { name: '!history @USER', value: 'Affiche l\'historique des services pour un utilisateur.', inline: true },
            { name: '!calcul', value: 'Calcule un total avec options de réduction.', inline: true }
        )
        .setColor(0x0000FF);
}

function isValidDate(dateString) {
    const regex = /^\d{4}-\d{2}-\d{2}$/;
    if (!regex.test(dateString)) return false;
    const date = new Date(dateString);
    return date instanceof Date && !isNaN(date);
}

client.on('messageCreate', async message => {
    if (message.author.bot) return;

    if (message.content.startsWith('!service')) {
        const userId = message.author.id;
        const member = message.guild.members.cache.get(userId);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const userStatus = await getUserStatus(displayName);

        const startButton = new ButtonBuilder()
            .setCustomId('startService')
            .setLabel('Débuter Service')
            .setStyle(ButtonStyle.Success)
            .setDisabled(userStatus === 'en service');

        const endButton = new ButtonBuilder()
            .setCustomId('endService')
            .setLabel('Terminer Service')
            .setStyle(ButtonStyle.Danger)
            .setDisabled(userStatus !== 'en service');

        const row = new ActionRowBuilder().addComponents(startButton, endButton);

        const sentMessage = await message.channel.send({
            content: 'Cliquez sur un bouton pour débuter ou terminer le service.',
            components: [row]
        });

        userMessageData[sentMessage.id] = { displayName, userId };
        await message.delete();
    } else if (message.content.startsWith('!temps')) {
        const mentionedUser = message.mentions.users.first();
        if (!mentionedUser) {
            await message.reply("Veuillez mentionner un utilisateur.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const today = new Date().toISOString().slice(0, 10);

        if (!isValidDate(today)) {
            await message.reply("La date fournie n'est pas valide.");
            return;
        }

        const totalTime = await getTotalTimeWorked(displayName, today);
        const embed = new EmbedBuilder()
            .setTitle(`Temps travaillé aujourd'hui pour ${displayName}`)
            .setDescription(`${displayName} a travaillé un total de ${totalTime} aujourd'hui.`)
            .setColor(0x00FF00);

        await message.reply({ embeds: [embed] });
        await message.delete();
    } else if (message.content.startsWith('!total')) {
        const args = message.content.split(' ');
        const numberOfDays = parseInt(args[1], 10) || 7;
        const mentionedUser = message.mentions.users.first();
        if (!mentionedUser) {
            await message.reply("Veuillez mentionner un utilisateur.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const endDate = new Date();
        const startDate = new Date(endDate);
        startDate.setDate(startDate.getDate() - numberOfDays);
        const formattedStartDate = startDate.toISOString().slice(0, 10);
        const formattedEndDate = endDate.toISOString().slice(0, 10);

        if (!isValidDate(formattedStartDate) || !isValidDate(formattedEndDate)) {
            await message.reply("Les dates fournies ne sont pas valides.");
            return;
        }

        const totalTime = await getTotalTimeWorkedInRange(displayName, formattedStartDate, formattedEndDate);
        const embed = new EmbedBuilder()
            .setTitle(`Temps travaillé pour ${displayName}`)
            .setDescription(`${displayName} a travaillé un total de ${totalTime} sur les ${numberOfDays} derniers jours.`)
            .setColor(0x00FF00);

        await message.reply({ embeds: [embed] });
        await message.delete();
    } else if (message.content.startsWith('!history')) {
        const mentionedUser = message.mentions.users.first();
        if (!mentionedUser) {
            await message.reply("Veuillez mentionner un utilisateur.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const history = await getServiceHistory(displayName);
        const embed = new EmbedBuilder()
            .setTitle(`Historique des services pour ${displayName}`)
            .setDescription(history)
            .setColor(0x00FF00);

        await message.reply({ embeds: [embed] });
        await message.delete();
    } else if (message.content === '!help') {
        const helpEmbed = createHelpEmbed();
        await message.reply({ embeds: [helpEmbed] });
    } else if (message.content === '!calcul') {
        await handleCalculation(message);
    }
});

client.on('interactionCreate', async interaction => {
    if (!interaction.isButton()) return;

    const { customId, user } = interaction;
    const messageId = interaction.message.id;
    const data = userMessageData[messageId];
    if (!data) return;

    const { displayName, userId } = data;
    const member = interaction.guild.members.cache.get(userId);
    const status = customId === 'startService' ? 'en service' : 'hors service';
    const timestamp = new Date();

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        let sheet = workbook.getWorksheet(displayName);
        if (!sheet) {
            sheet = workbook.addWorksheet(displayName);
            sheet.addRow(['Timestamp', 'Status']);
        }
        sheet.addRow([timestamp, status]);
        await workbook.xlsx.writeFile(excelFilePath);

        const serviceEmbed = createServiceEmbed(displayName, status, timestamp);
        await interaction.update({ embeds: [serviceEmbed], components: [] });
    } catch (error) {
        console.error('Erreur lors de l\'interaction avec le bouton:', error);
        await interaction.reply({ content: 'Une erreur est survenue lors de l\'interaction avec le bouton.', ephemeral: true });
    }
});

function sanitizeSheetName(name) {
    return name.replace(/[\[\]\*\/\\\?\:]/g, '_');
}

async function getTotalTimeWorked(displayName, date) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return '00:00:00';

        let totalTime = 0;
        let lastStartTime = null;

        sheet.eachRow({ includeEmpty: true }, (row) => {
            const timestampValue = row.getCell(1).value;
            const status = row.getCell(2).value;

            // Vérifier que le timestamp est valide
            let timestamp;
            if (timestampValue instanceof Date && !isNaN(timestampValue.getTime())) {
                timestamp = timestampValue;
            } else if (typeof timestampValue === 'string') {
                timestamp = new Date(timestampValue);
                if (isNaN(timestamp.getTime())) {
                    console.error('Valeur de timestamp invalide:', timestampValue);
                    return;
                }
            } else {
                console.error('Valeur de timestamp non reconnue:', timestampValue);
                return;
            }

            if (timestamp.toISOString().slice(0, 10) === date) {
                if (status === 'en service') {
                    lastStartTime = timestamp;
                } else if (status === 'hors service' && lastStartTime) {
                    totalTime += timestamp - lastStartTime;
                    lastStartTime = null;
                }
            }
        });

        if (lastStartTime) {
            totalTime += new Date() - lastStartTime;
        }

        return formatTime(totalTime);
    } catch (error) {
        console.error('Erreur lors de la récupération du temps travaillé:', error);
        return 'Erreur';
    }
}


async function getTotalTimeWorkedInRange(displayName, startDate, endDate) {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return '00:00:00';

        let totalTime = 0;
        let lastStartTime = null;

        sheet.eachRow({ includeEmpty: true }, (row) => {
            const timestampValue = row.getCell(1).value;
            const status = row.getCell(2).value;

            // Validation de la date
            let timestamp;
            if (timestampValue instanceof Date && !isNaN(timestampValue.getTime())) {
                timestamp = timestampValue;
            } else if (typeof timestampValue === 'string') {
                timestamp = new Date(timestampValue);
                if (isNaN(timestamp.getTime())) {
                    console.error('Valeur de timestamp invalide:', timestampValue);
                    return;
                }
            } else {
                console.error('Valeur de timestamp non reconnue:', timestampValue);
                return;
            }

            const timestampISODate = timestamp.toISOString().slice(0, 10);
            if (timestampISODate >= startDate && timestampISODate <= endDate) {
                if (status === 'en service') {
                    lastStartTime = timestamp;
                } else if (status === 'hors service' && lastStartTime) {
                    totalTime += timestamp - lastStartTime;
                    lastStartTime = null;
                }
            }
        });

        if (lastStartTime) {
            totalTime += new Date() - lastStartTime;
        }

        return formatTime(totalTime);
    } catch (error) {
        console.error('Erreur lors de la récupération du temps travaillé dans la plage:', error);
        return 'Erreur';
    }
}


function formatTime(milliseconds) {
    const totalSeconds = Math.floor(milliseconds / 1000);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

// Fonction pour formater les réponses de l'utilisateur
function formatCurrency(value) {
    return `$${value.toFixed(2)}`;
}

// Fonction pour lancer le calcul
async function handleCalculation(message) {
    let total = 0;

    const filter = response => response.author.id === message.author.id && !isNaN(response.content) && response.content.trim() !== '';
    const prompt = async (question, color) => {
        const embed = new EmbedBuilder()
            .setColor(color)
            .setDescription(question);
        await message.reply({ embeds: [embed] });
        const response = await message.channel.awaitMessages({ filter, max: 1, time: 60000, errors: ['time'] });
        return parseFloat(response.first().content);
    };

    const embed = new EmbedBuilder()
        .setTitle('Calcul du Total')
        .setDescription('Veuillez entrer les valeurs suivantes pour calculer le total.')
        .setColor(0x00FF00);

    let continueAdding = true;
    while (continueAdding) {
        const value = await prompt('Veuillez entrer une valeur :', 0x00FFFF);
        total += value * 1.25; // Ajouter 25%

        embed.setDescription(`Valeur ajoutée: ${formatCurrency(value)}.\nTotal actuel (avec 25% ajoutés) : ${formatCurrency(total)}`);
        await message.reply({ embeds: [embed] });

        const confirmationFilter = response => response.author.id === message.author.id && ['oui', 'non'].includes(response.content.toLowerCase());
        await message.reply({ embeds: [new EmbedBuilder().setColor(0xFFFF00).setDescription('Voulez-vous ajouter une autre valeur ? Répondez par "oui" ou "non".')] });
        const confirmation = await message.channel.awaitMessages({ filter: confirmationFilter, max: 1, time: 60000, errors: ['time'] });

        continueAdding = confirmation.first().content.toLowerCase() === 'oui';
    }

    embed.setDescription(`Le total avant réduction est : ${formatCurrency(total)}`);
    await message.reply({ embeds: [embed] });

    const discountFilter = response => response.author.id === message.author.id && ['oui', 'non'].includes(response.content.toLowerCase());
    await message.reply({ embeds: [new EmbedBuilder().setColor(0xFF0000).setDescription('Le client bénéficie-t-il d\'une réduction de 10% ? Répondez par "oui" ou "non".')] });
    const discountConfirmation = await message.channel.awaitMessages({ filter: discountFilter, max: 1, time: 60000, errors: ['time'] });

    if (discountConfirmation.first().content.toLowerCase() === 'oui') {
        total *= 0.90; // Retirer 10%
    }

    embed.setDescription(`Le prix total après réduction est : ${formatCurrency(total)}`);
    await message.reply({ embeds: [embed] });
}

client.login(process.env.TOKEN);