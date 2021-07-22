const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const credentials = JSON.parse(fs.readFileSync("./credentials.json", "utf8"));

const token = JSON.parse(fs.readFileSync("./token.json", "utf8"));

// Authorize a client with credentials, then call the Google Sheets API.
authorize(credentials, getValues);

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);
    oAuth2Client.setCredentials(token);
    callback(oAuth2Client);
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error while trying to retrieve access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

/**
 * Reading the document
 * @see https://docs.google.com/spreadsheets/d/1dHUF67GBDgI20NctSu3q6HvSdekfNnSbGe7bfGENoOg/edit#gid=0
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
function getValues(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    sheets.spreadsheets.values.get({
        spreadsheetId: '1dHUF67GBDgI20NctSu3q6HvSdekfNnSbGe7bfGENoOg',
        range: 'engenharia_de_software!A4:H',
    }, (err, res) => {
        if (err) return console.log('The API returned an error:' + err);
        const rows = res.data.values;
        if (rows.length) {
            console.log('reading the document');
            let values = rows.map((row) => {
                const faults = Number(row[2]) / 60;
                const grades = [Number(row[3]), Number(row[4]), Number(row[5])];
                const average = (grades[0] + grades[1] + grades[2]) / 3;
                const situation = getSituation(average, faults)
                let gradeFinal = 0;
                if (situation === "Exame Final") {
                    gradeFinal = getNAF(average);
                }
                return [situation, gradeFinal];
            });
            range = `!G4:H27`;
            //console.log(values);
            setValues(auth, values, range).then(() => {
                console.log("Update successful!")
            }).catch(error => {
                console.log(values);
                console.log(error)
            })
        } else {
            console.log('document not found');
        }
    });
}

function getSituation(average, faults) {
    if (faults > 0.25) {
        return "Reprovado por Falta";
    }
    else if (average < 50) {
        return "Reprovado por Nota";
    }
    else if (average >= 70) {
        return "Aprovado";
    }
    else {
        return "Exame Final"
    }
}

function getNAF(average) {
    return Math.ceil((100 - average));
}

//Writing in the document
function setValues(auth, values, range) {
    return new Promise((resolve, reject) => {
        //console.log('Atualizando a situação e a nota para aprovação final dos alunos');
        const resource = {
            values,
        };
        const sheets = google.sheets({ version: 'v4', auth });
        const spreadsheetId = '1dHUF67GBDgI20NctSu3q6HvSdekfNnSbGe7bfGENoOg'
        sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource,
        }, (err, result) => {
            if (err) {
                // Handle error
                console.log(resource)
                console.log(err);
                reject()
            } else {
                console.log('%d cells updated.', result.data.updatedCells);
                resolve()
            }
        });
    });
}