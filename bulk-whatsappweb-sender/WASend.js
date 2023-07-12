/** @format */

import wwebpkg from "whatsapp-web.js";
const { Client, LocalAuth, NoAuth, MessageMedia } = wwebpkg;

import qrcode from "qrcode-terminal";
import fs from "fs";

import parsePhoneNumber from "libphonenumber-js";

import mime from "mime-types";

import XlsxPopulate from "xlsx-populate";

// Function to add data to the Excel file
async function addDataToNotSent(variableData) {
  try {
    // Load the Excel file
    const workbook = await XlsxPopulate.fromFileAsync(
      "../BWWBS-GUI/NotSent.xlsx"
    );

    // Select the first sheet
    const sheet = workbook.sheet(0);

    // Find the last used row in column A
    const lastRow = sheet.usedRange().endCell("down").rowNumber();

    // Add the variable data to the next row
    sheet.cell(`A${lastRow + 1}`).value(variableData);

    // Save the modified workbook
    await workbook.toFileAsync("../BWWBS-GUI/NotSent.xlsx");
    console.log("Data added successfully to the Excel file.");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

async function addDataToSent(variableData) {
  try {
    // Load the Excel file
    const workbook = await XlsxPopulate.fromFileAsync("../BWWBS-GUI/Sent.xlsx");

    // Select the first sheet
    const sheet = workbook.sheet(0);

    // Find the last used row in column A
    const lastRow = sheet.usedRange().endCell("down").rowNumber();

    // Add the variable data to the next row
    sheet.cell(`A${lastRow + 1}`).value(variableData);

    // Save the modified workbook
    await workbook.toFileAsync("../BWWBS-GUI/Sent.xlsx");
    console.log("Data added successfully to the Excel file.");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}
// Usage
let variableData = "+919042";
// addDataToExcel(variableData);

import * as url from "url";
const __dirname = url.fileURLToPath(new URL(".", import.meta.url));

const log_file = fs.createWriteStream(__dirname + "/log.txt", { flags: "w" });

//Whatsapp magic
export default function sendMessages(numbersFile, messageToSend, mediaToSend) {
  const client = new Client({
    authStrategy: options.localAuth ? new LocalAuth() : new NoAuth(),
    puppeteer: {
      headless: process.env.NODE_ENV === "dev" ? false : true,
      executablePath: pupPath,
    },
  });

  client.initialize();

  client.on("qr", (qr) => {
    // NOTE: This event will not be fired if a session is specified.
    qrcode.generate(qr, { small: true });
  });

  client.on("auth_failure", (msg) => {
    // Fired if session restore was unsuccessful
    log("WWEB AUTHENTICATION FAILURE" + msg, true);
  });

  client.on("ready", async () => {
    log("WWEB READY");
    await client.sendPresenceAvailable();

    let my_number_country = parsePhoneNumber(
      "+" + client.info.wid.user
    ).country;

    let numbersArr;
    try {
      //const numbers = fs.readFileSync(path.resolve('numbers.txt'));
      const numbers = fs.readFileSync(numbersFile);
      numbersArr = numbers.toString().split(",");
    } catch (err) {
      log(err, true);
    }

    for (let number of numbersArr) {
      if (number !== "") {
        let parsed_number = parseNumber(number, my_number_country);
        if (parsed_number === null) {
          //if number is not valid skip it
          log(number + ": INVALID NUMBER", true);
          variableData = parsed_number;
          addDataToNotSent(variableData);
          continue;
        }

        await sendEverything(
          client,
          parsed_number + "@c.us",
          messageToSend,
          mediaToSend
        ); //'@c.us' represents a person's userdId

        //delay to try avoiding ban
        await new Promise((resolve) =>
          setTimeout(resolve, randBetween(delayms[0], delayms[1]))
        ); //in ms
        //await new Promise((resolve, reject) => setTimeout(resolve, 5000));	//in ms
      } //else nothing
    }
    await client.sendPresenceUnavailable();
    await client.destroy();
    log("ALL DONE!");
  });
}

async function sendEverything(WWebClient, chatId, messageToSend, mediaToSend) {
  //if number is not on Whatsapp
  if (!(await WWebClient.isRegisteredUser(chatId))) {
    log(chatId.split("@c.us")[0] + ": NOT ON WHATSAPP");
    variableData = chatId.split("@c.us")[0];
    addDataToNotSent(variableData);
  } else {
    let thisChat = await WWebClient.getChatById(chatId);

    //send seen
    await thisChat.sendSeen();

    //send "typing..."
    await thisChat.sendStateTyping();

    //if message exists
    if (messageToSend !== "") await thisChat.sendMessage(messageToSend);
    //if media exists
    for (let mediaPath of mediaToSend)
      if (mime.lookup(mediaPath).startsWith("video"))
        // This is necessary due to an issue with the library
        await thisChat.sendMessage(MessageMedia.fromFilePath(mediaPath), {
          sendMediaAsDocument: true,
        });
      else await thisChat.sendMessage(MessageMedia.fromFilePath(mediaPath));

    log(chatId.split("@c.us")[0] + ": SENT");
    variableData = chatId.split("@c.us")[0];
    addDataToSent(variableData);
  }
}

function parseNumber(number, country) {
  try {
    let parsed = parsePhoneNumber(number, country);
    return parsed.number.toString().replace(/[- )(+]/g, ""); //clean number
  } catch (err) {
    log(err, true);
    return null;
  }
}

function log(msg, error = false) {
  let today = new Date();
  let formattedDateTime =
    "[" +
    today.getFullYear() +
    "-" +
    (today.getMonth() + 1) +
    "-" +
    today.getDate() +
    " " +
    today.getHours() +
    ":" +
    today.getMinutes() +
    ":" +
    today.getSeconds() +
    "] ";

  if (!error) {
    console.log(msg);
    log_file.write(formattedDateTime + "INFO: " + msg + "\n");
  } else {
    console.error(msg);
    log_file.write(formattedDateTime + "ERROR: " + msg + "\n");
  }
}

function randBetween(min, max) {
  return Math.floor(Math.random() * (max - min + 1) + min);
}
