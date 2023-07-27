import { PublicClientApplication } from './msal-browser-2.14.2.js';

// const graphURL = 'https://graph.microsoft.com/v1.0';
// const baseURI = 'https://graph.microsoft.com/v1.0/sites/adobe.sharepoint.com,7be4993e-8502-4600-834d-2eac96f9558e,1f8af71f-8465-4c46-8185-b0a6ce9b3c85/drive/root:/theblog';

const graphURL = 'https://graph.microsoft.com/v1.0';
const baseURI = `https://graph.microsoft.com/v1.0/drives/b!9IXcorzxfUm_iSmlbQUd2rvx8XA-4zBAvR2Geq4Y2sZTr_1zgLOtRKRA81cvIhG1/root:/brandads`;
const driveId = 'b!9IXcorzxfUm_iSmlbQUd2rvx8XA-4zBAvR2Geq4Y2sZTr_1zgLOtRKRA81cvIhG1';
let connectAttempts = 0;
let accessToken;

const sp = {
    clientApp: {
        auth: {
            clientId: '2b4aa217-ddcd-4fe0-b09c-5a472764f552',
            authority: 'https://login.microsoftonline.com/fa7b1b5a-7b34-4387-94ae-d2c178decee1',
        },
    },
    login: {
        redirectUri: '/tools/sidekick/spauth.html',
    },
    api: {
        url: graphURL,
        file: {
            get: {
                baseURI,
            },
            download: {
                baseURI,
            },
            upload: {
                baseURI,
                method: 'PUT',
            },
            createUploadSession: {
                baseURI,
                method: 'POST',
                payload: {
                    '@microsoft.graph.conflictBehavior': 'replace',
                },
            },
        },
        directory: {
            create: {
                baseURI,
                method: 'PATCH',
                payload: {
                    folder: {},
                },
            },
        },
        driveUrl:baseURI,
        batch: {
            uri: `${graphURL}/$batch`,
        },
    },
};

export async function connect(callback) {
    console.log("I am in connect");
    const publicClientApplication = new PublicClientApplication(sp.clientApp);

    await publicClientApplication.loginPopup(sp.login);

    const account = publicClientApplication.getAllAccounts()[0];

    const accessTokenRequest = {
        scopes: ['files.readwrite', 'sites.readwrite.all'],
        account,
    };

    try {
        const res = await publicClientApplication.acquireTokenSilent(accessTokenRequest);
        accessToken = res.accessToken;
        if (callback) await callback();
    } catch (error) {
        // Acquire token silent failure, and send an interactive request
        if (error.name === 'InteractionRequiredAuthError') {
            try {
                const res = await publicClientApplication.acquireTokenPopup(accessTokenRequest);
                // Acquire token interactive success
                accessToken = res.accessToken;
                if (callback) await callback();
            } catch (err) {
                connectAttempts += 1;
                if (connectAttempts === 1) {
                    // Retry to connect once
                    connect(callback);
                }
                // Give up
                throw new Error(`Cannot connect to Sharepoint: ${err.message}`);
            }
        }
    }
}

function validateConnnection() {
    if (!accessToken) {
        throw new Error('You need to sign-in first');
    }
}

function getRequestOption() {
    validateConnnection();

    const bearer = `Bearer ${accessToken}`;
    const headers = new Headers();
    headers.append('Authorization', bearer);

    return {
        method: 'GET',
        headers,
    };
}

// async function createFolder(folder) {
//     validateConnnection();
//
//     const options = getRequestOption();
//     options.headers.append('Accept', 'application/json');
//     options.headers.append('Content-Type', 'application/json');
//     options.method = sp.api.directory.create.method;
//     options.body = JSON.stringify(sp.api.directory.create.payload);
//
//     const res = await fetch(`${sp.api.directory.create.baseURI}${folder}`, options);
//     if (res.ok) {
//         return res.json();
//     }
//     throw new Error(`Could not create folder: ${folder}`);
// }


async function createFolder(folderPath) {
    validateConnnection();

    const options = getRequestOption();
    options.headers.append('Accept', 'application/json');
    options.headers.append('Content-Type', 'application/json');
    options.method = sp.api.directory.create.method;
    options.body = JSON.stringify(sp.api.directory.create.payload);


    const res = await fetch(`${sp.api.directory.create.baseURI}/${folderPath}`, options);

    if (res.ok) {
        return res.json();
    } else if (res.status === 409) {
        // Folder already exists, return the existing folder
        return getFolder(folderPath);
    }

    throw new Error(`Could not create or get folder: ${folderPath}`);
}

async function getFolder(folderPath) {
    validateConnnection();

    const options = getRequestOption();
    options.method = 'GET';

    const res = await fetch(`${sp.api.driveUrl}/${folderPath}`, options);
    if (res.ok) {
        return res.json();
    }

    throw new Error(`Could not get folder: ${folderPath}`);
}


// async function createExcelFile(filePath) {
//     validateConnection();
//
//     const options = getRequestOption();
//     options.method = 'GET';
//
//     const res = await fetch(`${sp.api.url}/drives/${driveId}/root:${filePath}`, options);
//
//     if (res.ok) {
//         // File already exists, return the existing file
//         return res.json();
//     } else if (res.status === 404) {
//         // File not found, create a new Excel file
//         return createNewExcelFile(filePath);
//     }
//
//     throw new Error(`Could not check or create Excel file: ${filePath}`);
// }
//
async function createNewExcelFile() {
    validateConnnection();
    const folderPath = '/ad3';
    const filename = 'match.xlsx';

    const options = getRequestOption();
    options.headers.append('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    options.method = 'PUT';
    options.body = JSON.stringify({
        '@microsoft.graph.conflictBehavior': 'replace',
        name: filename,
        file: {},
    });

    const res = await fetch(`${baseURI}${folderPath}/${filename}:/content`, options);

    if (res.ok) {
        return res.json();
    }

    throw new Error(`Could not create Excel file: ${filePath}`);
}

async function createExcelFile(folderPath, fileName) {
    validateConnnection();

    const options = getRequestOption();
    options.headers.append('Content-Type', 'application/json');
    options.method = 'PUT';
    options.body = JSON.stringify({});

    const res = await fetch(`${baseURI}${folderPath}/${fileName}:/workbook`, options);
    if (res.ok) {
        return true;
    } else {
        throw new Error(`Could not create Excel file ${fileName} in folder ${folderPath}`);
    }
}


export async function checkAndUpdateExcelFile() {
    validateConnnection();

    const response = await getDriveId1();
    //const response = downloadUploadDocumentOnSite();



//     const folderPath = '/ad3';
//     const filename = 'match.xlsx';
//     const docName = `sampledoc.docx`;
//
//     const sheetName = 'defaultsheet';
//     //const searchText = 'push notification';
//     const entry = {
//         id: 'abc2',
//         notify: 'event2',
//         sent: 'yes2'
//     };
//     const entries = [{
//         id: 'abc3',
//         notify: 'event3',
//         sent: 'yes3'
//     },
//         {
//             id: 'abc4',
//             notify: 'event4',
//             sent: 'yes4'
//         },
//     ];
//
//     //const fileId = await getFileId(folderPath,filename);
//     const documentId = await getFileId(folderPath,docName);
//     //const sheetId = await getSheetId(fileId, sheetName);
//     //await addEntriesToExcel(fileId, sheetName, entries);
//     const siteId = await getSiteId();
//     //const dataResponse = await updateDocument(siteId, documentId);
//
//
//     var documentUrl = window.location.href;
//     console.log(documentUrl);
// // Make an HTTP GET request to retrieve the document content
//
//     var xhr = new XMLHttpRequest();
//     xhr.open("GET", documentUrl, true);
//     xhr.setRequestHeader("Authorization", `Bearer ${accessToken}`);
//     xhr.onreadystatechange = function() {
//         if (xhr.readyState === 4 && xhr.status === 200) {
//             // Document content is available in the response
//             var documentContent = xhr.responseText;
//             console.log(documentContent);
//             // Further processing of the document content
//         }
//     };
//     xhr.send();
//
//
//    // const rewriteResponse = await rewriteDocument(siteId, documentId);
//     //const searchResponse = await searchdocument(siteId, documentId);
// //    const downlodedFile = await downloadUploadDocument(siteId, documentId);
//     //await findText(fileId, sheetName, entry);
//     //await findTextInExcel(fileId, sheetName, entry.id);
//      //await createFolder(folderPath);
//      //await createNewExcelFile();
//     //
//     // // Check if the sheet exists
//     //const sheetExists = await doesSheetExist(folderPath, filename, sheetName);
//     // if (!sheetExists) {
//     //     // Create the sheet if it does not exist
//     //     await createSheet(folderPath, filename, sheetName);
//     // }
//     //
//     // // Find the row index containing the search text
//     // const rowIndex = await findRowIndex(`${folderPath}/${filename}`, sheetName, searchText);
//     //
//     // // Add the new entry below the found row or at the end
//     // await addNotificationEntry(`${folderPath}/${filename}`, sheetName, searchText, entry);
}

async function downloadUploadDocument(sitesid, documentid) {
    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/items/${documentid}/content`;

    validateConnnection();

    const updateContent = 'New document content';
    const options = getRequestOption();
    options.method='GET';
    options.headers.append('Content-Type', 'application/json');
    // options.body = JSON.stringify({
    //     content: updateContent,
    // });


    const response = await fetch(`${endpoint}`, options);

    if (response.ok) {
        const blob = await response.blob();
        const file = new File([blob], 'document.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        // const downloadLink = document.createElement('a');
        // downloadLink.href = URL.createObjectURL(file);
        // downloadLink.download = file.name;
        //
        // // Trigger the download
        // downloadLink.click();
        //
        // // Clean up the temporary URL object
        // URL.revokeObjectURL(downloadLink.href);
        //
        // console.log('Downloaded document:', file);

        //upload document
        const options1 = getRequestOption();
        options1.method='PUT';
        options1.body=file;
        const response1 = await fetch(`https://graph.microsoft.com/v1.0/sites/${sitesid}/drives/${driveId}/root:/brandads/ad3/sample1.docx:/content`, options1);
        if (response1.ok) {
            console.log('Document uploaded successfully');
        }

        return blob;
    }

    throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
}




async function downloadUploadDocumentOnSite(sitesid, documentid) {

// API endpoint
    try {
        //get document id
        const siteId = "FirstSite";

        // File path within the SharePoint site
        const filePath = "first/First.docx";

        validateConnnection();

        const options = getRequestOption();
        options.method = 'GET';
        options.headers.append('Content-Type', 'application/json');

        // options.method='GET';
        // options.

        //adobe.sharepoint.com,d7196e56-4780-4a65-9249-504609568d95,b3d702f4-e849-4da5-84bc-616aa7f5ab17"


        const url = `https://graph.microsoft.com/v1.0/sites/adobe.sharepoint.com:/sites/FirstSite?$select=id`

        // const siteUrl = "https://adobe.sharepoint.com/sites/FirstSite/Shared%20Documents/Forms/AllItems.aspx";
        // const url = `https://graph.microsoft.com/v1.0/sites?filter=webUrl eq '${siteUrl}'`;
        const response = await fetch(url, options);
       // if (response.ok) {
            const responseData = await response.json();
            console.log(responseData);
      //  }

    } catch (error) {
        console.log(error);
    }

}







async function rewriteDocument(sitesid, documentid) {
    const updatedContent = 'This is the updated content of the document.';
    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/items/${documentid}/content`;

    validateConnnection();

    const options = getRequestOption();
    options.method='PATCH';
    options.headers.append('Content-Type', 'text/plain');
    options.body = updatedContent


    const response = await fetch(`${endpoint}`, options);

    if (response.ok) {
        const data = await response.json();
        return data;
    }

    throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
}
async function updateDocument(sitesid, documentid) {
    const endpoint = `https://graph.microsoft.com/v1.0/sites/${sitesid}/drive/items/${documentid}/content`;

    validateConnnection();

    const updateContent = 'New document content';
    const options = getRequestOption();
    options.method='PATCH';
    options.redirect='manual';
    options.headers.append('Content-Type', 'application/json');
    options.headers.append('Origin',window.location.origin);
    options.body = JSON.stringify({
        content: updateContent,
    });


    const response = await fetch(`${endpoint}`, options);

    if (response.redirected) {
        const redirectedUrl = response.url;
//        const data = await response.json();
        console.log(redirectedUrl);
        return data;
    }

    throw new Error(`Could not add entries to doc file. Status: ${response.status}`);
}

async function searchdocument(sitesid, documentid) {
    const searchQuery = 'PandaBin';
    const endpoint = `https://graph.microsoft.com/v1.0/sites/${sitesid}/drive/root/search(q='${encodeURIComponent(searchQuery)}')`;

    validateConnnection();

    const updateContent = 'New document content';
    const options = getRequestOption();
    options.method='GET';
    options.headers.append('Content-Type', 'application/json');
    // options.body = JSON.stringify({
    //     requests: [
    //         {
    //             entityTypes: ['driveItem'],
    //             query: {
    //                 queryString: searchQuery,
    //             },
    //         },
    //     ],
    // });


    const response = await fetch(`${endpoint}`, options);

    if (response.ok) {
        const data = await response.json();
        return data;
    }

    throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
}

async function addEntriesToExcel(fileId, sheetName, entries) {
    const endpoint = `/drives/${driveId}/items/${fileId}/workbook/worksheets('${sheetName}')/range(address='A3:C4')`;

    // const requestBody = {
    //     values: [[entries.id, entries.notify, entries.sent]],
    // };

    const requestBody = {
        values: entries.map((entry) => [entry.id, entry.notify, entry.sent]),
    };

    validateConnnection();

    const options = getRequestOption();
    options.method='PATCH';
    options.headers.append('Content-Type', 'application/json');
    options.body = JSON.stringify(requestBody);


    const response = await fetch(`${graphURL}${endpoint}`, options);

    if (response.ok) {
        return response.json();
    }

    throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
}

async function findText(fileId, sheetName, entries) {
    const endpoint = `/drives/${driveId}/items/${fileId}/workbook/worksheets('${sheetName}')/range(address='A1:C2')`;

    // const requestBody = {
    //     values: [[entries.id, entries.notify, entries.sent]],
    // };

    validateConnnection();

    const options = getRequestOption();
    options.method='GET';
    options.headers.append('Content-Type', 'application/json');
    //options.body = JSON.stringify(requestBody);


    const response = await fetch(`${graphURL}${endpoint}`, options);

    if (response.ok) {
        const searchResults = await response.json();

        // Find text in the 2D array
        const searchText = 'abc2';
        let found = false;
        let rowIndex, columnIndex;

        for (let row = 0; row < searchResults.values.length; row++) {
            //for (let col = 0; col < searchResults.values[row].length; col++) {
                if (searchResults.values[row][0] === searchText) {
                    rowIndex = row + 1; // Adding 1 to row and column indices since Excel starts from 1
                    columnIndex = 0 + 1;
                    found = true;
                    break;
                }
            //}
            if (found) {
                break;
            }
        }

        if (found) {
            console.log(`Text '${searchText}' found at Row: ${rowIndex}, Column: ${columnIndex}`);
        } else {
            console.log(`Text '${searchText}' not found in the array.`);
        }
        return;
    }

    throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
}
async function getFileId(folderPath, fileName) {
    const endpoint = `${sp.api.directory.create.baseURI}${folderPath}/${fileName}`;

    validateConnnection();

    const options = getRequestOption();
    options.headers.append('Content-Type', 'application/json');
    options.method = 'GET';

    const response = await fetch(`${endpoint}`, options);

    if (response.ok) {
        const file = await response.json();
        return file.id;
    }

    throw new Error(`Could not retrieve file ID. Status: ${response.status}`);
}

async function getSiteId() {
    const endpoint = `${sp.api.directory.create.baseURI}`;

    validateConnnection();

    const options = getRequestOption();
    options.headers.append('Content-Type', 'application/json');
    options.method = 'GET';

    const response = await fetch(`${endpoint}`, options);

    if (response.ok) {
        const data  = await response.json();
        return data.parentReference.siteId;
    }

    throw new Error(`Could not retrieve file ID. Status: ${response.status}`);
}


async function getSheetId(workbookId, sheetName) {
//    const endpoint = `${sp.api.directory.create.baseURI}${folderPath}/${fileName}`;
    const endpoint =`/drives/${driveId}/items/${workbookId}/workbook/worksheets`;

    validateConnnection();

    const options = getRequestOption();
    options.headers.append('Content-Type', 'application/json');
    options.method = 'GET';

    const response = await fetch(`${graphURL}${endpoint}`, options);

    if (response.ok) {
        const worksheets = await response.json();
        const worksheet = worksheets.value.find(w => w.name === sheetName);
        if (!worksheet) {
            console.error(`Worksheet "${sheetName}" not found.`);
            return;
        }
        return worksheet.id;
    }

    throw new Error(`Could not retrieve file ID. Status: ${response.status}`);
}


// async function findTextInExcel(fileId, sheetName, searchText) {
// //    const endpoint = `/drives/${driveId}/items/${fileId}/workbook/worksheets('${sheetName}')/search(q='${encodeURIComponent(searchText)}')`;
//     var endpointUrl = 'https://graph.microsoft.com/v1.0/me/drive/items/' + fileId + '/workbook/worksheets/' + sheetName + '/range(address=' + sheetName + '!A1:Z1000)?$search="' + searchText + '"';
//
//     validateConnnection();
//
//     const options = getRequestOption();
//     options.method='GET';
//     options.headers.append('Content-Type', 'application/json');
//
//
// //    const response = await fetch(`${graphURL}${endpoint}`, options);
//     const response = await fetch(`${endpointUrl}`,options);
//     if (response.ok) {
//         const searchResults = await response.json();


//
//         return { rowNumber, columnNumber };
//     }
//
//     throw new Error(`Could not find the specified text. Status: ${response.status}`);
// }

async function findTextInExcel(fileId,sheetName, searchText) {
    //const endpointUrl = `/drives/${driveId}/items/${fileId}/workbook/worksheets/00000000-0001-0000-0000-000000000000/usedRange/search(q='${searchText}')`;
    //const endpointUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheetName}')/usedRange/find(values="${encodeURIComponent(searchText)}")`;

    //const range = `${sheetName}!A1:C2`;
    const endpointUrl = `/drive/items/${fileId}/workbook/worksheets('${sheetName}')/range(address='A1:C2')?$expand=values`;

    validateConnnection();
    const options = getRequestOption();
    options.method='GET';
    options.headers.append('Content-Type', 'application/json');

    const response = await fetch(`${graphURL}${endpointUrl}`, options);
    if (response.ok) {
        const searchResults = await response.json();
        const firstResult = searchResults.value[0]; // Assuming there's at least one match

        // Retrieve the row and column numbers of the first match
        const rowNumber = firstResult.row;
        const columnNumber = firstResult.column;

        return { rowNumber, columnNumber };
    }

    throw new Error(`Could not find the specified text. Status: ${response.status}`);
}

export async function getDriveId() {
    const publicClientApplication = new PublicClientApplication(sp.clientApp);

    await publicClientApplication.loginPopup();

    const account = publicClientApplication.getAllAccounts()[0];

    const accessTokenRequest = {
        scopes: ['files.readwrite', 'sites.readwrite.all'],
        account,
    };

    try {
        const res = await publicClientApplication.acquireTokenSilent(accessTokenRequest);
        const accessToken = res.accessToken;

        const options = {
            headers: {
                Authorization: `Bearer ${accessToken}`,
            },
        };

        const driveResponse = await fetch('https://graph.microsoft.com/v1.0/me/drive', options);
        const driveData = await driveResponse.json();
        const driveId = driveData.id;

        return driveId;
    } catch (error) {
        throw new Error('Failed to retrieve drive ID');
    }
}


async function getDriveId1() {
    try {
        validateConnnection();
        console.log(accessToken);
        const driveResponse = await fetch('https://graph.microsoft.com/v1.0/drives', getRequestOption());
        const driveData = await driveResponse.json();
        const driveId = driveData.id;

        return driveId;
    } catch (error) {
        throw new Error('Failed to retrieve drive ID');
    }
}

export async function saveFile(file, dest) {
    validateConnnection();

    const folder = dest.substring(0, dest.lastIndexOf('/'));
    const filename = dest.split('/').pop().split('/').pop();

    await createFolder(folder);

    // start upload session

    const payload = {
        ...sp.api.file.createUploadSession.payload,
        description: 'Preview file',
        fileSize: file.size,
        name: filename,
    };

    let options = getRequestOption();
    options.headers.append('Accept', 'application/json');
    options.headers.append('Content-Type', 'application/json');
    options.method = sp.api.file.createUploadSession.method;
    options.body = JSON.stringify(payload);

    let res = await fetch(`${sp.api.file.createUploadSession.baseURI}${dest}:/createUploadSession`, options);
    if (res.ok) {
        const json = await res.json();

        options = getRequestOption();
        // TODO API is limited to 60Mb, for more, we need to batch the upload.
        options.headers.append('Content-Length', file.size);
        options.headers.append('Content-Range', `bytes 0-${file.size - 1}/${file.size}`);
        options.method = sp.api.file.upload.method;
        options.body = file;

        res = await fetch(`${json.uploadUrl}`, options);
        if (res.ok) {
            return res.json();
        }
    }
    throw new Error(`Could not upload file ${dest}`);
}

async function findRowIndex(sheetName, searchText) {
    validateConnnection();

    const options = getRequestOption();
    options.method = 'GET';

    const res = await fetch(`${sp.api.file.get.baseURI}${sheetName}:/workbook/tables/Table1/rows`, options);
    if (res.ok) {
        const json = await res.json();
        const rows = json.value;

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const cellValues = row.values.map(cell => cell.text);

            if (cellValues.includes(searchText)) {
                return i + 1; // Add 1 to convert from zero-based index to one-based index (Excel row index starts from 1)
            }
        }
    }
    return -1; // Return -1 if the search text is not found in any row
}

async function addNotificationEntry(sheetName, searchText, entry) {
    validateConnnection();

    const rowIndex = await findRowIndex(sheetName, searchText);
    const insertRowIndex = rowIndex !== -1 ? rowIndex + 1 : -1; // Insert row below the found row or at the end

    const options = getRequestOption();
    options.headers.append('Accept', 'application/json');
    options.headers.append('Content-Type', 'application/json');
    options.method = 'POST';
    options.body = JSON.stringify({
        index: insertRowIndex,
        values: [
            [entry.id, entry.notify, entry.sent]
        ]
    });

    const res = await fetch(`${sp.api.file.get.baseURI}${sheetName}:/workbook/tables/Table1/rows/add`, options);
    if (res.ok) {
        return res.json();
    }
    throw new Error(`Could not add notification entry to the sheet: ${sheetName}`);
}

export async function test() {
    validateConnnection();

    const options = getRequestOption();
    options.headers.append('Accept', 'application/json');
    options.headers.append('Content-Type', 'application/json');
    options.method = 'GET';
    // options.body = JSON.stringify(payload);

    await fetch(`${sp.api.file.createUploadSession.baseURI}`, options);
    throw new Error('Could not upload file');
}

// async function addEntriesToExcel(fileId, sheetName, entries) {
//     console.log("in add entries");
//
//     const endpoint = `/drives/${driveId}/items/${fileId}/workbook/worksheets('${sheetName}')/range(address='A6:C7')`;
//
//     // const requestBody = {
//     //     values: [[entries.id, entries.notify, entries.sent]],
//     // };
//
//     const requestBody = {
//         values: entries.map((entry) => [entry.id, entry.notify, entry.sent]),
//     };
//
//     validateConnnection();
//
//     const options = getRequestOption();
//     options.method='PATCH';
//     options.headers.append('Content-Type', 'application/json');
//     options.body = JSON.stringify(requestBody);
//
//
//     const response = await fetch(`${graphURL}${endpoint}`, options);
//
//     if (response.ok) {
//         console.log("entries updated");
//         return "updated";
//     }
//
//     throw new Error(`Could not add entries to Excel file. Status: ${response.status}`);
// }
//
// async function getFileId(driveid, folderPath, fileName) {
//     const endpoint = `${graphURL}/drives/${driveId}/root:${folderPath}/${fileName}`;
//     'https://graph.microsoft.com/v1.0';
//     `https://graph.microsoft.com/v1.0/drives/b!9IXcorzxfUm_iSmlbQUd2rvx8XA-4zBAvR2Geq4Y2sZTr_1zgLOtRKRA81cvIhG1/root:/brandads`;
//     // const endpoint = `${sp.api.directory.create.baseURI}${folderPath}/${fileName}`;
//
//     validateConnnection();
//
//     const options = getRequestOption();
//     options.headers.append('Content-Type', 'application/json');
//     options.method = 'GET';
//
//     const response = await fetch(`${endpoint}`, options);
//
//     if (response.ok) {
//         const file = await response.json();
//         return file.id;
//     }
//
//     throw new Error(`Could not retrieve file ID. Status: ${response.status}`);
// }

// async function getDriveId() {
//
//     validateConnnection();
//
//     const options = getRequestOption();
//     options.method = 'GET';
//
//     try{
//     const response = await fetch('`https://graph.microsoft.com/v1.0/sites/AdobeFranklinPOC/drives', options);
//         if (response.ok) {
//             const data = await response.json();
//             driveId = data.value[0].id;
//             console.log('Drive ID:', driveId);
//         }
//     } catch (error) {
//         throw new Error('Failed to retrieve drive ID');
//     }
// }