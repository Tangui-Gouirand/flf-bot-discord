const { Client, GatewayIntentBits, ButtonBuilder, ActionRowBuilder, ButtonStyle, EmbedBuilder, StringSelectMenuBuilder } = require('discord.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const moment = require('moment');

const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent, GatewayIntentBits.GuildMembers] });
const excelFilePath = path.join(__dirname, 'services.xlsx');
const userMessageIds = {}; // Pour stocker les messages avec les boutons

// Fonction pour créer un fichier Excel vide s'il n'existe pas
async function createExcelFileIfNotExists() {
    try {
        if (!fs.existsSync(excelFilePath)) {
            console.log(`Le fichier Excel n'existe pas. Création de ${excelFilePath}`);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.writeFile(excelFilePath);
            console.log(`Fichier Excel créé à l'emplacement : ${excelFilePath}`);
        } else {
            console.log(`Le fichier Excel existe déjà à l'emplacement : ${excelFilePath}`);
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

// Fonction pour obtenir le statut de l'utilisateur
async function getUserStatus(displayName) {
    try {
        console.log(`Lecture du fichier Excel à l'emplacement : ${excelFilePath}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);

        if (!sheet) return 'hors service';

        const lastRow = sheet.lastRow;
        if (!lastRow) return 'hors service';

        const status = lastRow.getCell(2).value;
        return status || 'hors service';
    } catch (error) {
        console.error('Erreur lors de la récupération du statut de l\'utilisateur:', error);
        return 'Erreur';
    }
}

// Fonction pour obtenir l'historique des services
async function getServiceHistory(displayName) {
    try {
        console.log(`Lecture du fichier Excel à l'emplacement : ${excelFilePath}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return 'Aucun historique trouvé.';

        let history = '';
        sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const [timestamp, status] = [row.getCell(1).value, row.getCell(2).value];
            history += `\n${formatDate(new Date(timestamp))} - ${status}`;
        });
        return history || 'Aucun historique trouvé.';
    } catch (error) {
        console.error('Erreur lors de la récupération de l\'historique des services:', error);
        return 'Erreur';
    }
}

// Fonction pour formater la date et l'heure
function formatDate(date) {
    return moment(date).format('dddd, MMMM Do YYYY, h:mm:ss a');
}

// Fonction pour créer un message de service
function createServiceEmbed(displayName, status, timestamp) {
    const color = status === 'en service' ? 0x00FF00 : 0xFF0000; // Vert pour débuté, rouge pour terminé
    const action = status === 'en service' ? 'débuté' : 'terminé';
    return new EmbedBuilder()
        .setColor(color)
        .setDescription(`Service ${action} pour ${displayName} à ${formatDate(timestamp)}.`);
}

// Fonction pour ajouter un menu déroulant
function createUserSelectMenu() {
    const userSelectMenu = new StringSelectMenuBuilder()
        .setCustomId('userSelect')
        .setPlaceholder('Sélectionnez un utilisateur')
        .addOptions([
            // Ajoutez ici les options pour les utilisateurs
            { label: 'Utilisateur 1', value: 'user1' },
            { label: 'Utilisateur 2', value: 'user2' }
        ]);

    return new ActionRowBuilder().addComponents(userSelectMenu);
}

// Fonction pour créer un message d'aide
function createHelpEmbed() {
    return new EmbedBuilder()
        .setTitle('Commandes du Bot')
        .addFields(
            { name: '!service', value: 'Commence ou termine un service.', inline: true },
            { name: '!temps @USER', value: 'Affiche le temps travaillé aujourd\'hui pour un utilisateur.', inline: true },
            { name: '!total [N] @USER', value: 'Affiche le temps travaillé pour un utilisateur sur les derniers jours.', inline: true },
            { name: '!history @USER', value: 'Affiche l\'historique des services pour un utilisateur.', inline: true }
        )
        .setColor(0x0000FF);
}

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
            .setStyle(ButtonStyle.Success)
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
        const mentionedUser = message.mentions.users.first();

        if (!mentionedUser) {
            await message.reply("Veuillez mentionner un utilisateur.");
            return;
        }

        const member = message.guild.members.cache.get(mentionedUser.id);
        const displayName = member ? sanitizeSheetName(member.displayName) : 'Unknown User';
        const today = new Date().toISOString().slice(0, 10); // Date d'aujourd'hui au format YYYY-MM-DD

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

        if (!isValidDate(startDate) || !isValidDate(endDate)) {
            await message.reply("Les dates fournies ne sont pas valides.");
            return;
        }

        const totalTime = await getTotalTimeWorkedInRange(displayName, startDate.toISOString().slice(0, 10), endDate.toISOString().slice(0, 10));

        const embed = new EmbedBuilder()
            .setTitle(`Temps travaillé pour ${displayName}`)
            .setDescription(`${displayName} a travaillé un total de ${totalTime} durant les ${numberOfDays} derniers jours.`)
            .setColor(0x00FF00);

        await message.reply({ embeds: [embed] });

        // Supprimer le message d'origine
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

        // Supprimer le message d'origine
        await message.delete();
    } else if (message.content === '!help') {
        const helpEmbed = createHelpEmbed();
        await message.reply({ embeds: [helpEmbed] });
    }
});

// Fonction pour gérer les interactions avec les boutons
client.on('interactionCreate', async interaction => {
    if (!interaction.isButton() && !interaction.isSelectMenu()) return;

    try {
        if (interaction.isButton()) {
            const userMessage = userMessageIds[interaction.message.id];
            if (!userMessage) return;

            const { displayName, channelId } = userMessage;
            const status = interaction.customId === 'startService' ? 'en service' : 'hors service';
            const timestamp = new Date();

            // Enregistrer les données dans le fichier Excel
            console.log(`Écriture dans le fichier Excel à l'emplacement : ${excelFilePath}`);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(excelFilePath);
            let sheet = workbook.getWorksheet(displayName);

            if (!sheet) {
                sheet = workbook.addWorksheet(displayName);
                sheet.addRow(['Timestamp', 'Status']);
                console.log(`Nouvelle feuille créée pour ${displayName}`);
            }

            sheet.addRow([timestamp.toISOString(), status]);
            await workbook.xlsx.writeFile(excelFilePath);
            console.log(`Données enregistrées pour ${displayName} à ${timestamp.toISOString()}`);

            // Créer un embed avec la couleur appropriée
            const embed = createServiceEmbed(displayName, status, timestamp);

            // Répondre à l'interaction avec l'embed
            await interaction.reply({ embeds: [embed], ephemeral: true });

            // Mettre à jour le message avec les boutons
            const startButton = new ButtonBuilder()
                .setCustomId('startService')
                .setLabel('Début de Service')
                .setStyle(ButtonStyle.Success)
                .setDisabled(status === 'en service');

            const endButton = new ButtonBuilder()
                .setCustomId('endService')
                .setLabel('Fin de Service')
                .setStyle(ButtonStyle.Danger)
                .setDisabled(status !== 'en service');

            const row = new ActionRowBuilder().addComponents(startButton, endButton);
            await interaction.message.edit({ components: [row] });
        } else if (interaction.isSelectMenu()) {
            // Gérer les sélections d'utilisateurs si nécessaire
        }
    } catch (error) {
        console.error('Erreur lors de l\'interaction :', error);
        await interaction.reply('Une erreur est survenue.');
    }
});

// Fonction pour obtenir le temps total travaillé pour un utilisateur sur une journée
async function getTotalTimeWorked(displayName, date) {
    try {
        console.log(`Lecture du fichier Excel à l'emplacement : ${excelFilePath}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return '0h 0m 0s';

        let totalSeconds = 0;
        let lastStartTime = null;

        sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const [timestamp, status] = [row.getCell(1).value, row.getCell(2).value];
            const rowDate = new Date(timestamp).toISOString().slice(0, 10);

            if (rowDate === date) {
                if (status === 'en service') {
                    lastStartTime = new Date(timestamp);
                } else if (status === 'hors service' && lastStartTime) {
                    const endTime = new Date(timestamp);
                    totalSeconds += (endTime - lastStartTime) / 1000;
                    lastStartTime = null;
                }
            }
        });

        return formatTime(totalSeconds);
    } catch (error) {
        console.error('Erreur lors de la récupération du temps total travaillé:', error);
        return 'Erreur';
    }
}

// Fonction pour obtenir le temps total travaillé dans une plage de dates
async function getTotalTimeWorkedInRange(displayName, startDate, endDate) {
    try {
        console.log(`Lecture du fichier Excel à l'emplacement : ${excelFilePath}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelFilePath);
        const sheet = workbook.getWorksheet(displayName);
        if (!sheet) return '0h 0m 0s';

        let totalSeconds = 0;
        let lastStartTime = null;

        sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const [timestamp, status] = [row.getCell(1).value, row.getCell(2).value];
            const rowDate = new Date(timestamp).toISOString().slice(0, 10);

            if (rowDate >= startDate && rowDate <= endDate) {
                if (status === 'en service') {
                    lastStartTime = new Date(timestamp);
                } else if (status === 'hors service' && lastStartTime) {
                    const endTime = new Date(timestamp);
                    totalSeconds += (endTime - lastStartTime) / 1000;
                    lastStartTime = null;
                }
            }
        });

        return formatTime(totalSeconds);
    } catch (error) {
        console.error('Erreur lors de la récupération du temps total travaillé dans la plage de dates:', error);
        return 'Erreur';
    }
}

// Fonction pour formater le temps en heures, minutes et secondes
function formatTime(seconds) {
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const secs = Math.round(seconds % 60); // Arrondir les secondes à l'entier le plus proche
    return `${hours}h ${minutes}m ${secs}s`;
}

// Fonction pour assainir le nom d'une feuille Excel
function sanitizeSheetName(name) {
    return name.replace(/[\/\\?*[\]]/g, '_').substring(0, 31);
}

client.login(process.env.TOKEN);
