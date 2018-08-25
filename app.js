const Oxford = require('oxford-dictionary');
const readline = require('readline');
const Excel = require('exceljs');
const rl = readline.createInterface(process.stdin, process.stdout);

// dotenv is being used to protect API credentials. To obtain an API key for Oxford Dictionaries, visit https://developer.oxforddictionaries.com/?tag=#plans.
require('dotenv').config();

const config = {
    app_id: process.env.APP_ID,
    app_key: process.env.APP_KEY,
    source_lang: "en"
};

// Create an instance of the Oxford dictionary.
const dict = new Oxford(config);

console.log("Welcome to the Lexical Amplifier Node app!");

// Create an array to store a session's queried words and their definitions.
const queriesList = [];

async function lookup(word) {
    try {
        // create an array for each query containing the queried word and its definitions
        let entry = [
            word
        ];

        let firstStep = await dict.definitions(word); // dict.definitions(word) returns a promise for the word's dictionary results, which is a JSON containing nested objects and arrays which are looped through below to isolate the definitions
    
        let jsonDepth1 = firstStep['results'][0]['lexicalEntries'];
        for(let i=0; i < jsonDepth1.length; i++) {
            let jsonDepth2 = jsonDepth1[i]['entries'][0]['senses'];
            for(let j = 0; j < jsonDepth2.length; j++) {
                let jsonDepth3 = jsonDepth2[j]['definitions'];
                for(let k = 0; k < jsonDepth3.length; k++) {
                    let definition = jsonDepth3[k].trim();
                    console.log("- " + definition);

                    entry.push(definition);
                }
            }
        }

        queriesList.push(entry);
    } catch(err) {
        console.log("Word not found!");
    }
}

// Build the question-answer chain.
function askUser() {
    rl.question("Please enter the word you would like to define: ", async function(word) {
        await lookup(word);
        rl.question("Would you like to define another word? Type 'yes' if so or press any key to quit: ", (askAnother) => {
            if (askAnother.toLowerCase().trim() === 'yes') {
                askUser();
            } else {
                rl.question("Would you like a full list of the queried words and their definitions before exiting? Type 'yes' if so or press any key to quit: ", (answer) => {
                    if (answer.toLowerCase().trim() === 'yes') {
                        rl.question("Would you like to export the list as an Excel (.xlsx) file as well? Type 'yes' if so or press any key to simply print a list to the screen before exiting: ", (excelAnswer) => {
                            if (excelAnswer.toLowerCase().trim() === 'yes') {
                                console.log(queriesList);
                                generateSpreadsheet();
                                console.log("Your Excel file has been downloaded to the same folder that this Node app resides in.");
                                rl.close();
                            } else {
                                console.log(queriesList);
                                rl.close();
                            }
                        })
                    } else {
                        rl.close();
                    }
                })
            }
        })
    })
}

// Create the Excel file.
function generateSpreadsheet() {
    try {
        const workbook = new Excel.Workbook();

        workbook.creator = 'Lexical Amplifier';
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("Definitions");

        worksheet.columns = [
            { header: 'Word', key: 'word', width: 36 },
            { header: 'Definition(s)', key: 'def', width: 120 }
        ];

        worksheet.getCell('A1').font = {
            bold: true
        }
        worksheet.getCell('A1').alignment = {
            horizontal: 'center'
        }
        worksheet.getCell('A1').fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{ argb:'FFFFFF00' },
            bgColor:{ argb:'FFFFFF00' }
        }

        worksheet.getCell('B1').font = {
            bold: true,
            color: { argb: 'FFFFFFFF'}
        }
        worksheet.getCell('B1').alignment = {
            horizontal: 'center'
        }
        worksheet.getCell('B1').fill = {
            type: 'pattern',
            pattern:'solid',
            fgColor:{ argb:'FF002060' },
            bgColor:{ argb:'FF002060' }
        }

        worksheet.addRows(queriesList);

        workbook.xlsx.writeFile('Lexical_Amplifier_Results.xlsx');
    } catch(err) {
        console.log(err);
    }
}

// Run the app by initiating the question-answer chain.
askUser();