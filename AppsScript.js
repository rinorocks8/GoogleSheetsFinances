function main() {
  const token = getAuthToken();

  const accounts = fetchAccountDetails(token);
  insertAccountsIntoSheet(accounts);

  const accountIDs = getAccountIdsToPull();
  const allTransactions = getTransactions(token, accountIDs);
  insertTransactionsIntoSheet(allTransactions);
  formatTransactions();
}

function getAuthToken() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings");
  const KeyRange = sheet.getRange("A5");
  const DateRange = sheet.getRange("A7");

  if (Date.now() - DateRange.getValue() < 72000000) {
    console.log("Using Cached Login!");
    return KeyRange.getValue();
  }

  console.log("Login Init");
  initiateLogin();

  const attempts = [5000, 15000, 30000];

  for (let i = 0; i < attempts.length; i++) {
    Utilities.sleep(attempts[i]);
    console.log("Login Check (Attempt " + (i + 1) + ")");
    try {
      return authenticateWithCode();
    } catch (error) {
      console.log(
        "Failed to authenticate (Attempt " + (i + 1) + "): " + error.message
      );

      // If it's the last attempt, throw the error
      if (i === attempts.length - 1) throw error;
    }
  }
}

function storeToken(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Settings");
  const KeyRange = sheet.getRange("A5");
  KeyRange.setValue(token);
  const DateRange = sheet.getRange("A7");
  DateRange.setValue(Date.now());
}

function initiateLogin() {
  var url = "https://auth.quiltt.io/v1/users/session";
  var payload = {
    session: {
      deploymentId: "8501cbef-e303-421f-946e-5a1e9e4b0a71",
      email: Session.getEffectiveUser().getEmail(),
    },
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(url, options);
}

function fetchCodeFromEmail() {
  var threads = GmailApp.search(
    'from:support@quiltt.io subject:"🗝 Your Quiltt Hub passcode" is:unread'
  );

  if (threads.length == 0) throw new Error("No relevant email threads found.");

  // Assuming the latest thread is the first one in the search results. Get the latest message from the latest thread.
  var messages = threads[0].getMessages();
  var latestMessage = messages[messages.length - 1];
  var body = latestMessage.getPlainBody();

  // Extracting code using a regular expression.
  var match = body.match(/Your Quiltt Hub passcode is \*(\d{6})\*/);
  //Delete email
  threads[0].moveToTrash();

  if (match) {
    return match[1];
  } else {
    throw new Error("Login code not found in email.");
  }
}

function authenticateWithCode() {
  var code = fetchCodeFromEmail();

  var url = "https://auth.quiltt.io/v1/users/session";
  var payload = {
    session: {
      deploymentId: "8501cbef-e303-421f-946e-5a1e9e4b0a71",
      email: Session.getEffectiveUser().getEmail(),
      passcode: code,
    },
  };

  var options = {
    method: "put",
    contentType: "application/json",
    payload: JSON.stringify(payload),
  };

  var response = UrlFetchApp.fetch(url, options);
  const token = JSON.parse(response.getContentText()).token;
  storeToken(token);
  return token;
}

function postGraphQLRequest(authToken, query) {
  // Get the token first
  var url = "https://api.quiltt.io/v1/graphql";
  var headers = {
    Authorization: "Bearer " + authToken,
    "Content-Type": "application/json",
  };

  var payload = {
    operationName: "SpendingAccountsWithTransactionsQuery",
    query: query,
  };

  var options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, // This will prevent the script from stopping if there's an HTTP exception
  };

  var response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() == 200) {
    return JSON.parse(response.getContentText()); // This contains the response from the server
  } else {
    // Handle errors as needed
    Logger.log("Error: " + response.getContentText());
    return null;
  }
}

function getTransactions(token, accountIDs) {
  let endCursor = null;
  let allTransactions = [];

  do {
    const result = postGraphQLRequest(
      token,
      `
      query SpendingAccountsWithTransactionsQuery {
        transactionsConnection(
          filter: {accountIds: ${JSON.stringify(accountIDs)}}
          sort: DATE_DESC
          after: "${endCursor || ""}"
        ) {
          nodes {
            id
            amount
            date
            description
            account {
              name
              id
            }
            source(type: PLAID) {
              ... on PlaidTransaction {
                checkNumber
              }
            }
          }
          pageInfo {
            endCursor
          }
        }
      }
    `
    );

    if (result && result.data && result.data.transactionsConnection) {
      const transactions = result.data.transactionsConnection.nodes;
      allTransactions.push(...transactions);

      endCursor = result.data.transactionsConnection.pageInfo.endCursor;
    } else {
      break;
    }
  } while (endCursor);
  return allTransactions;
}

function fetchAccountDetails(token) {
  const result = postGraphQLRequest(
    token,
    `
    query SpendingAccountsWithTransactionsQuery {
      connections {
        accounts {
          name
          id
          institution {
            name
          }
          balance {
            current
          }
        }
      }
    }
  `
  );
  if (result && result.data && result.data.connections) {
    console.log("Fetched All Account Data");
    return result.data.connections
      .map((connection) => connection.accounts)
      .flat();
  } else {
    Logger.log("Error: Unexpected response format or no Accounts found.");
  }
}

function insertTransactionsIntoSheet(allTransactions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Transactions");
  let AutoCategorySheet = ss.getSheetByName("AutoCategory");

  // If the sheet doesn't exist, create one and set the headers.
  const headers = [
    "ID",
    "Account_ID",
    "Date",
    "Description",
    "Category",
    "Sub-Category",
    "Amount",
  ];
  if (!sheet) {
    sheet = ss.insertSheet("Transactions");
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  // Create a hash set for all IDs in the sheet for quick lookup.
  const lastRow = sheet.getLastRow() + 1;
  const idsInSheet = new Set(
    sheet
      .getRange(2, 1, lastRow - 1)
      .getValues()
      .flat()
      .map(String)
  );

  const AutoCategoryInSheet = new Map(
    AutoCategorySheet.getRange(2, 1, AutoCategorySheet.getLastRow() - 1, 3)
      .getValues()
      .filter((row) => row[0] !== "")
      .map((row) => [
        row[0],
        {
          category: row[1],
          subcategory: row[2],
        },
      ])
  );

  const rowsToInsert = [];
  allTransactions.forEach((transaction) => {
    if (!idsInSheet.has(transaction.id)) {
      // If ID not in sheet
      let category = "",
        subCategory = "";
      for (const [pattern, categorization] of AutoCategoryInSheet.entries()) {
        const regex = new RegExp(pattern, "i"); // assuming case-insensitive match
        if (regex.test(transaction.description)) {
          category = categorization.category;
          subCategory = categorization.subcategory;
          break; // Break once a match is found
        }
      }

      const row = [
        transaction.id,
        transaction.account.id,
        transaction.date,
        transaction.description,
        category,
        subCategory,
        transaction.amount,
        // transaction?.source?.checkNumber || "",
      ];
      rowsToInsert.push(row);
    }
  });

  // Insert rows in bulk at the top below headers
  if (rowsToInsert.length > 0) {
    sheet.insertRowsAfter(1, rowsToInsert.length); // Insert new rows below header
    sheet
      .getRange(2, 1, rowsToInsert.length, rowsToInsert[0].length)
      .setValues(rowsToInsert);
  }
}

function insertAccountsIntoSheet(accounts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Accounts");

  // If the sheet doesn't exist, create one and set the headers.
  if (!sheet) {
    sheet = ss.insertSheet("Accounts");
    const headers = ["ID", "Institution", "Name", "Balance", "Pull?"];
    sheet.appendRow(headers);
  }

  // Create a map for all IDs in the sheet to their row numbers for quick lookup.
  const lastRow = sheet.getLastRow() + 1;
  const idsInSheet = new Map();
  const idValues = sheet.getRange(2, 1, lastRow - 1).getValues();

  for (let i = 0; i < idValues.length; i++) {
    idsInSheet.set(idValues[i][0], i + 2); // +2 because sheets are 1-indexed and we have a header row
  }

  accounts.forEach((account) => {
    const row = [
      account.id,
      account.institution.name,
      account.name,
      account.balance.current,
    ];
    if (idsInSheet.has(account.id)) {
      // If ID is in sheet, update the row
      const rowNumber = idsInSheet.get(account.id);
      sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    } else {
      // Else, append a new row
      sheet.appendRow(row);
    }
  });
}

function getAccountIdsToPull() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Accounts");

  if (!sheet) {
    throw new Error("Accounts sheet not found!");
  }

  const lastRow = sheet.getLastRow();
  const pullColumn = 5; // Assuming "Pull?" is the 5th column
  const idColumn = 1; // Assuming "id" is the 1st column

  // Fetching entire columns for "id" and "Pull?"
  const ids = sheet.getRange(2, idColumn, lastRow - 1).getValues();
  const pullValues = sheet.getRange(2, pullColumn, lastRow - 1).getValues();
  const idsToPull = [];
  for (let i = 0; i < pullValues.length; i++) {
    if (pullValues[i][0] === "Yes") {
      idsToPull.push(ids[i][0]);
    }
  }
  return idsToPull;
}

function formatTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const TransactionsSheet = ss.getSheetByName("Transactions");
  const AutoCategorySheet = ss.getSheetByName("AutoCategory");
  const SettingsSheet = ss.getSheetByName("Settings");

  var range = TransactionsSheet.getRange("A2:G");
  range.sort([
    { column: 3, ascending: false },
    { column: 7, ascending: false },
  ]);

  const lastRow = TransactionsSheet.getLastRow();
  TransactionsSheet.getRange(1, 7, lastRow).setNumberFormat(
    `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`
  );
  TransactionsSheet.getRange(1, 3, lastRow).setNumberFormat("MM/dd/yy");

  // Data validation for Category
  const categoryRange = SettingsSheet.getRange("B2:B");
  const categoryValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(categoryRange, true)
    .build();
  TransactionsSheet.getRange("E2:E").setDataValidation(categoryValidation);
  AutoCategorySheet.getRange("B2:B").setDataValidation(categoryValidation);

  // Copy the data validation rule from the source range to the target range
  // Only way to do it without static references
  const Transaction_sourceRange = SettingsSheet.getRange("A2");
  const transactionRange = TransactionsSheet.getRange("F2:F");
  Transaction_sourceRange.copyTo(
    transactionRange,
    SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    false
  );

  const AutoCategory_sourceRange = SettingsSheet.getRange("A3");
  const AutoCategoryRange = AutoCategorySheet.getRange("C2:C");
  AutoCategory_sourceRange.copyTo(
    AutoCategoryRange,
    SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
    false
  );
}

function getSubCategory(categoryRange, SettingsCategoryRange) {
  const sub_categories = {};
  SettingsCategoryRange.forEach((item) => {
    sub_categories[item[0]] = item.slice(1).filter(Boolean);
  });
  return [...categoryRange].map((category) =>
    category[0].length !== 0 ? sub_categories[category[0]] : ""
  );
}
