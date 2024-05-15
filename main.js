const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
require("dotenv").config({ path: "./.env" });
const { exec } = require("child_process");
const sudo = require("sudo-prompt");
//const remote = require('electron').remote;
const express = require("express");
const fs = require("fs");
const bodyParser = require("body-parser");
const pipedrive = require("pipedrive");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const expressApp = express();
expressApp.use(bodyParser.json());
const proposalDirectory = "Path";
const wordDocumentPath = "Template Doc Path";

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 350,
    height: 600,
    webPreferences: {
      //preload: path.join(__dirname, 'preload.js'), // Specify the preload script
      nodeIntegration: true, // Enable Node.js integration
      contextIsolation: false, // Allow context isolation to access ipcRenderer
    },
  });

  // Load your HTML file
  mainWindow.loadFile(path.join(__dirname, "index.html"));

  ipcMain.on("get-documents", (event) => {
    // Provide a path to your asset folder containing documents
    const sharedFolderPath = path.join(__dirname, "assets");

    // Read the contents of the asset folder
    fs.readdir(sharedFolderPath, (err, files) => {
      if (err) {
        console.error(err);
        event.sender.send("documents", []);
      } else {
        console.log("List of documents:", files);
        event.sender.send("documents", files);
      }
    });
  });
}

ipcMain.on("btn-click", () => {
  console.log("Button clicked!");
});

ipcMain.on("inputField", (event, term) => {
  console.log(`Received term from renderer process: ${term}`);
  // You can use the 'term' value for further processing
  // Validate the 'term' (5-digit number)
  if (/^\d{5}$/.test(term)) {
    let matchingFolderPath = findMatchingFolder(proposalDirectory, term);
    if (matchingFolderPath != null) {
      event.sender.send(
        "inputField-validation",
        "Matching folder found - creating document in: " + matchingFolderPath,
      );

      main(term);
    } else {
      event.sender.send(
        "inputField-validation",
        "No matching folder found. Attempting to create new folder...",
      );
      main(term);
    }
  } else {
    event.sender.send(
      "inputField-validation",
      "Invalid input. Please enter 5 digits.",
    );
  }
});

function findMatchingFolder(rootDir, searchTerm) {
  const folders = fs.readdirSync(rootDir);

  for (const folder of folders) {
    // Check if the folder name starts with a 5-digit number
    const folderPath = path.join(rootDir, folder);

    if (fs.statSync(folderPath).isDirectory() && /^\d{5}/.test(folder)) {
      //console.log('Checking folder:', folderPath);
      const folderName = folder.trim().substring(0, 5);
      //console.log('comparing', folderName, 'with', searchTerm, '...');
      const temp = String(searchTerm).trim().substring(0, 5);
      if (folderName === temp) {
        console.log("Match found:", folderName);
        return folderPath;
      }
    }
  }

  return null; // No matching folder found
}
function updateTemplate(docContent, matchingFolderPath) {
  //const documentPath = "BKL-Template-Letterhead-for-scripting.docx";
  const documentPath = wordDocumentPath;
  // Load the docx file as binary content
  const content = fs.readFileSync(
    path.resolve(__dirname, documentPath),
    "binary",
  );

  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });
  // Perform the template replacement
  doc.render({
    title: docContent[0] || "Title Not Found",
    organizationName: docContent[1] || "Organization Name Not Found",
    personName: docContent[2] || "Person Not Found",
    organizationAddress: docContent[3] || "Organization Address Not Found",
    proposalNumber: docContent[4] || "Proposal Number Not Found",
    firstName: docContent[5] || "First Name Not Found",
    ownerName: docContent[6] || "Owner Name Not Found",
    ownerEmail: docContent[7] || "Owner Email Not Found",
    picName: docContent[8] || "PIC Name Not Found",
    picEmail: docContent[9] || "PIC Email Not Found",
  });

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
  });
  const newFolderPath = path.join(matchingFolderPath, "20 Doc Prep"); // Specify the new folder path

  const newDocumentName =
    String(docContent[4]) +
    " " +
    String(docContent[0]) +
    " - R0.docx"; // Specify the new document name

  // Write the filled document to the new path
  fs.writeFileSync(path.resolve(newFolderPath, newDocumentName), buf);
  
  // Now, the filled document is saved in the directory
  console.log("Document filled and saved successfully.");
}
async function getDocContent(proposalNumber) {
  try {
    const opts = {
      exactMatch: true,
    };
    const defaultClient = new pipedrive.ApiClient();
    defaultClient.authentications.api_key.apiKey =
      process.env.PIPEDRIVE_API_KEY;
    const api = new pipedrive.DealsApi(defaultClient);

    // First API call
    const data = await api.searchDeals(proposalNumber, opts);
    if (data.data.items.length === 0) {
      throw new Error("No deals found");
      return null;
    }
    //display deal data
    console.log("Deal data:");
    console.log("data:", data);

    //save deals data in a json file
    fs.writeFileSync("data.json", JSON.stringify(data, null, 2));

    const dealItem = data.data.items[0].item;

    const title = dealItem.title || "Title Not Found";
    const organizationName =
      dealItem.organization?.name || "Organization Name Not Found";
    const person = dealItem.person?.name || "Person Not Found";
    let organizationAddress =
      dealItem.organization?.address || "Organization Address Not Found";
    organizationAddress = splitAddress(organizationAddress);

    const firstName = person.split(" ")[0];

    // Second API call
    let opts2 = {
      start: 0,
      limit: 56,
    };

    const data3 = await api.getDeal(dealItem.id);
    fs.writeFileSync("data3.json", JSON.stringify(data3, null, 2));
    const ownerData =
      data3.data?.b8fae9b957863db254f44b872b6ab7b2dd3534d6 ||
      "Owner Name Not Found";

    const ownerName = ownerData?.name || "Owner Name Not Found";
    const ownerEmail = ownerData?.email || "Owner Email Not Found";

    const PicData = data3.data?.e0b557f66c631fa886554a894621dce70cd8b57b;
    const PicName = PicData?.name || "PIC Name Not Found";
    const PicEmail = PicData?.email || "PIC Email Not Found";

    const docContent = [
      title,
      organizationName,
      person,
      organizationAddress,
      proposalNumber,
      firstName,
      ownerName,
      ownerEmail,
      PicName,
      PicEmail,
    ];

    return docContent;
  } catch (error) {
    console.error("Error:", error);
    throw new Error("Error fetching doc content");
  }
}
function createProposalFolder(proposalNumber, title) {
  const folderPath = path.join(
    "Path",
    `${proposalNumber} ${title}`,
  );
  const inFolderPath = path.join(folderPath, "10 In");
  const docFolderPath = path.join(folderPath, "20 Doc Prep");
  const issuedFolderPath = path.join(folderPath, "25 Issued");

  // Check if the folder already exists
  if (!fs.existsSync(folderPath)) {
    // Create the folder
    fs.mkdirSync(folderPath);
    console.log(`Folder created: ${folderPath}`);
  } else {
    console.log(`Folder already exists: ${folderPath}`);
  }
  if (!fs.existsSync(inFolderPath)) {
    // Create the folder
    fs.mkdirSync(inFolderPath);
    console.log(`Folder created: ${inFolderPath}`);
  } else {
    console.log(`Folder already exists: ${inFolderPath}`);
  }
  if (!fs.existsSync(docFolderPath)) {
    // Create the folder
    fs.mkdirSync(docFolderPath);
    console.log(`Folder created: ${docFolderPath}`);
  } else {
    console.log(`Folder already exists: ${docFolderPath}`);
  }
  if (!fs.existsSync(issuedFolderPath)) {
    // Create the folder
    fs.mkdirSync(issuedFolderPath);
    console.log(`Folder created: ${issuedFolderPath}`);
  } else {
    console.log(`Folder already exists: ${issuedFolderPath}`);
  }
}
function copyFile(destinationPath) {
  // Check if the source file exists
  const sourcePath = wordDocumentPath;

  if (!fs.existsSync(sourcePath)) {
    console.error("Source file does not exist");
    return;
  }

  // Check if the destination folder exists, create if not
  if (!fs.existsSync(path.dirname(destinationPath))) {
    fs.mkdirSync(path.dirname(destinationPath), { recursive: true });
  }

  fs.copyFile(sourcePath, destinationPath, (err) => {
    if (err) {
      console.error("Error:", err);
      return;
    }
    console.log("File copied successfully");
  });
}
//Since the address format is unpredicatble, we need to clean it by splitting it into components
function splitAddress(address) {
  let normalizedAddress = address.replace(/[,]/g, "").trim("");

  const cityPattern =
    /Abbotsford|Armstrong|Burnaby|Campbell River|Castlegar|Chilliwack|Colwood|Coquitlam|Courtenay|Cranbrook|Dawson Creek|Delta|Duncan|Enderby|Fernie|Fort St. John|Grand Forks|Greenwood|Kamloops|Kelowna|Kimberley|Langford|Langley|Maple Ridge|Merritt|Mission|Nanaimo|Nelson|New Westminster|North Vancouver|Parksville|Penticton|Pitt Meadows|Port Alberni|Port Coquitlam|Port Moody|Powell River|Prince George|Prince Rupert|Quesnel|Revelstoke|Richmond|Rossland|Salmon Arm|Surrey|Terrace|Trail|Vancouver|Vernon|Victoria|West Kelowna|White Rock|Williams Lake/i;
  const postalCodePattern = /[A-Z]\d[A-Z] \d[A-Z]\d|[a-z]\d[a-z] \d[a-z]\d/i;
  const countryPattern = /Canada|USA|United States|Mexico/i;
  const provincePattern = /BC/i;
  const postalCodeMatch = normalizedAddress.match(postalCodePattern);
  const provinceMatch = normalizedAddress.match(provincePattern);
  const countryMatch = normalizedAddress.match(countryPattern);
  const cityMatch = normalizedAddress.match(cityPattern);
  let postalCode = postalCodeMatch ? postalCodeMatch[0] : null;
  let province = provinceMatch ? provinceMatch[0] : null;
  let country = countryMatch ? countryMatch[0] : null;
  let city = cityMatch ? cityMatch[0] : null;
  let componentsRemoved = normalizedAddress;
  if (postalCode) {
    componentsRemoved = componentsRemoved.replace(postalCode, "").trim(", ");
  }
  if (province) {
    componentsRemoved = componentsRemoved.replace(province, "").trim(", ");
  }
  if (country) {
    componentsRemoved = componentsRemoved.replace(country, "").trim(", ");
  }
  if (city) {
    componentsRemoved = componentsRemoved.replace(city, "").trim(", ");
  }
  let addressLine2 = city + ", " + province + ", " + country + " " + postalCode;
  let addressLines = componentsRemoved + "\n" + addressLine2;
  addressLines = removeNull(addressLines);
  return addressLines;
}
//Not working, need to find a different way to access outlook folders
function runExecutable(args) {
  return new Promise((resolve, reject) => {
    //const executablePath = path.join(__dirname, 'createOutlookFolder.exe');
    const executablePath = "./createOutlookFolder.exe";
    console.log("Args:", args);
    // Ensure args are properly quoted
    const command = `"${executablePath}" "${args.join('" "')}"`;

    console.log(`Executing command: ${command}`);

    sudo.exec(
      command,
      { name: "BKLProposalGenerator" },
      (error, stdout, stderr) => {
        if (error) reject(new Error(`Execution Failed: ${error.message}`));
        else if (stderr)
          reject(new Error(`Executable Error Output: ${stderr}`));
        else resolve(stdout.trim());
      },
    );

    // Optional: Log real-time output from the executable
  });
}
function removeNull(address) {
  let cleanedString = address.replace(/null/g, "").trim();
  cleanedString = removeOrphanCommas(cleanedString);

  return cleanedString;
}
function removeOrphanCommas(address) {
  let cleanedString = address.replace(/,,/g, "").trim();
  cleanedString = address.replace(/, ,/g, "").trim();
  return cleanedString;
}
async function main(term) {
  const proposalNumber = term;
  try {
    const proposalDirectory = "Path";

    if (proposalNumber != null) {
      //setProposalNumberForDeal(dealId, proposalNumber);
      let matchingFolderPath = findMatchingFolder(
        proposalDirectory,
        proposalNumber,
      );
      try {
        const docContent = await getDocContent(proposalNumber);

        console.log("docContent:", docContent);
        const title = docContent[0];
        console.log("title:", title);
        let outlookFolderName = proposalNumber + " " + title;
        let arguments = [outlookFolderName];

        if (matchingFolderPath) {
          updateTemplate(docContent, matchingFolderPath);
          /* Commented out until solution is found for running outlook management script
                                  runExecutable(arguments)
                                      .then((output) => {
                                      console.log('PowerShell script executed successfully:', output);
                                      })
                                      .catch((error) => {
                                      console.error('Failed to execute PowerShell script:', error);
                                      });
                                      */
        } else {
          console.log("No matching folder found");
          createProposalFolder(proposalNumber, title);
          matchingFolderPath = findMatchingFolder(
            proposalDirectory,
            proposalNumber,
          );
          updateTemplate(docContent, matchingFolderPath);
          /* 
                                  runExecutable(arguments)
                                      .then((output) => {
                                      console.log('PowerShell script executed successfully:', output);
                                      })
                                      .catch((error) => {
                                      console.error('Failed to execute PowerShell script:', error);
                                      });
                                      */
        }
      } catch (error) {
        console.error("Error:", error.message);
        // If doc can't be generated, move the template to the '20 Doc Prep' folder
        matchingFolderPath =
          matchingFolderPath +
          "Path";
        copyFile(matchingFolderPath);
      }
    }
  } catch (error) {
    console.log("Error:", error);
  }
}

app.whenReady().then(createWindow);

//------------------------------------EXTRA FUNCTIONS --------------------------------------------------
// function readDocumentContents(document) {
//   return new Promise((resolve, reject) => {
//     // Specify the path to the documents directory
//     const documentsDirectory = path.join(__dirname, 'assets');
//     const filePath = path.join(documentsDirectory, document);

//     mammoth.extractRawText({ path: filePath })
//       .then((result) => {
//         const text = result.value; // The raw text
//         let output = '';
//         var textLines = text.split("\n");

//         for (var i = 0; i < textLines.length; i++) {
//           output += textLines[i] + '\n';
//         }
//       // rosolve(text) == formatted text resolve(output) == raw text
//         resolve(text);
//       })
//       .catch((error) => {
//         console.error(error);
//         reject(error);
//       });
//   });
// }


// function makeAxiosRequest() {
// // Make a GET request to your server
// axios.get(serverURL)
//     .then(response => {
//         const axiosData = response.axiosData;
//         console.log(axiosData); // Do something with the data
//     })
//     .catch(error => {
//         console.error('Error fetching data from the server:', error);
//     });
// }
