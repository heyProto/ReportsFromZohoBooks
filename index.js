const axios = require("axios");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const args = require("args");

args.option("token", "Access token to be used for pulling data");
args.option("orgid", "Organization ID for Zoho Books account");
args.option("projectid", "Project ID for generating report");
args.option("attachment", "Choose to download attachments for report (y or n)");
args.option(
  "exchangerate",
  "[Optional] USD rate to be used. Defaults to uncalculated if not mentioned."
);

const flags = args.parse(process.argv);
if (!flags.token || !flags.orgid || !flags.projectid || !flags.attachment) {
  console.error("Invalid usage");
  args.showHelp();
  process.exit(1);
}

const authToken = flags.token;
const orgId = flags.orgid;
const projectId = flags.projectid;
const allowAttachment = flags.attachment;
const exchangeRate = flags.exchangerate;
let isDownloadable = false;
let project_name = "Default-" + projectId;

const dollarFormat = "$0.00";
const rupeeFormat = "₹0.00";

let apiCallCount = 0;
let vendorMap = new Map();

if (allowAttachment) {
  switch (allowAttachment.toLowerCase()) {
    case "y":
      isDownloadable = true;
      break;
    case "n":
      isDownloadable = false;
      break;
    default:
      console.error("Invalid parameter. Expected (y/n)");
      process.exit(1);
  }
} else {
  console.error("Invalid parameter. Expected (y/n)");
  process.exit(1);
}

let axiosInstance = axios.create({
  baseURL: "https://books.zoho.in/api/v3/",
  headers: {
    Authorization: "Zoho-authtoken " + authToken,
  },
  timeout: 300000,
});

axiosInstance.interceptors.request.use(request => {
  console.log('Starting request to URL: "', request.url, '"');
  apiCallCount++;
  return request;
});

let projectHeader = [];
let projectFooter = [];

axiosInstance
  .get("projects/" + projectId, {
    organization_id: orgId,
  })
  .then(res => {
    console.log("Found project. Accumulating project details...");
    projectHeader = buildProjectHeader(res.data.project);
    projectFooter = buildProjectFooter();
    project_name = res.data.project.project_name;
    findDetails();
  })
  .catch(err => {
    console.error("Unable to find project ID: ", projectId);
    console.error("Error while retrieving project: ", err);
    fetchProjects(1).then(projects => {
      console.error("Did you mean any of the following? ");
      projects.forEach(project => {
        console.error(project.project_id, ": ", project.project_name);
      });

      process.exit(1);
    });
  });

function findDetails() {
  axios
    .all([fetchExpenses(1), fetchBills(1)])
    .then(
      axios.spread((expenses, bills) => {
        let expenseData = expenses;
        let mappedExpense = expenseData.map(element => {
          let expenseId = element.expense_id;
          let commonExpense = {
            "Transaction Date": element.date,
            // Type: 'Expense',
            // 'Reference Number': expenseId,
            Notes: element.attachment_name
              ? element.reference_number + " - " + element.attachment_name
              : element.reference_number,
            // Status: element.status,
            // Vendor: '-'
          };
          if (element.has_attachment) {
            commonExpense["Attachment"] =
              "https://books.zoho.in/api/v3/expenses/" +
              expenseId +
              "/attachment";
            if (isDownloadable) {
              downloadAttachment(
                "expenses/" + expenseId + "/attachment",
                "expenses",
                commonExpense["Attachment"]
              );
            }
          } else {
            commonExpense["Attachment"] = "Not available";
          }

          return loadExpenses(expenseId, commonExpense);
        });
        let billData = bills;
        let mappedBills = billData.map(element => {
          let billId = element.bill_id;
          let commonBill = {
            "Transaction Date": element.date,
            // Type: 'Bill',
            // 'Reference Number': billId,
            Notes: element.bill_number + " - " + element.attachment_name,
            // Status: element.status
          };

          if (element.has_attachment) {
            commonBill["Attachment"] =
              "https://books.zoho.in/api/v3/bills/" + billId + "/attachment";
            if (isDownloadable) {
              downloadAttachment(
                "bills/" + billId + "/attachment",
                "bills",
                commonBill["Attachment"]
              );
            }
          } else {
            commonBill["Attachment"] = "Not available";
          }

          return loadBills(billId, commonBill, element.vendor_id);
        });

        Promise.all([...mappedExpense, ...mappedBills])
          .then(x => {
            let y = x.reduce((a, b) => a.concat(b), []);
            let inr_column = "D";
            let exrate_column = "E";
            let subtotal_row = y.length + 8;

            y = y.map((entry, i) => {
              return {
                "Transaction Date": entry["Transaction Date"],
                "Description of Good/Service": entry["Description"],
                Category: entry["Category"],
                "Amount in INR": {
                  f: entry["Amount in INR"],
                  t: "n",
                  z: rupeeFormat,
                },
                "Exchange Rate": entry["Exchange Rate"],
                "Amount in USD": {
                  f: entry["Amount in USD"],
                  t: "n",
                  z: dollarFormat,
                },
                Notes: entry["Notes"],
                "Attachment URL": entry["Attachment"],
              };
            });
            let ws2 = xlsx.utils.aoa_to_sheet(projectHeader);
            xlsx.utils.sheet_add_json(ws2, y, { origin: "A6" });

            xlsx.utils.sheet_add_aoa(
              ws2,
              [
                [""],
                [
                  "",
                  "",
                  "Subtotal",
                  {
                    f:
                      "SUM(" +
                      inr_column +
                      "7:" +
                      inr_column +
                      (y.length + 6) +
                      ")",
                    t: "n",
                    z: rupeeFormat,
                  },
                ],
                [
                  "",
                  "",
                  "PROTO Fee",
                  {
                    f: "0.1*" + inr_column + subtotal_row,
                    t: "n",
                    z: rupeeFormat,
                  },
                ],
                [
                  "",
                  "",
                  "Subtotal with fee",
                  {
                    f:
                      inr_column +
                      subtotal_row +
                      "+" +
                      inr_column +
                      (subtotal_row + 1),
                    t: "n",
                    z: rupeeFormat,
                  },
                ],
                [
                  "",
                  "",
                  "GST",
                  {
                    f: "0.18*" + inr_column + (subtotal_row + 2),
                    t: "n",
                    z: rupeeFormat,
                  },
                ],
                [
                  "",
                  "",
                  "Total",
                  {
                    f:
                      inr_column +
                      (subtotal_row + 2) +
                      "+" +
                      inr_column +
                      (subtotal_row + 3),
                    t: "n",
                    z: rupeeFormat,
                  },
                  { f: exchangeRate || 1, t: "n" },
                  {
                    f:
                      inr_column +
                      (subtotal_row + 4) +
                      "/" +
                      exrate_column +
                      (subtotal_row + 4),
                    t: "n",
                    z: dollarFormat,
                  },
                ],
              ],
              { origin: -1 }
            );

            xlsx.utils.sheet_add_aoa(ws2, projectFooter, { origin: -1 });

            let ws1 = createFinancialReport();
            let ws3 = createAdvanceRequestForm();

            let wb = xlsx.utils.book_new();
            xlsx.utils.book_append_sheet(wb, ws1, "Financial Report");
            xlsx.utils.book_append_sheet(wb, ws2, "Expense Report");
            xlsx.utils.book_append_sheet(wb, ws3, "Advance Request Form");

            xlsx.writeFile(wb, project_name + ".xlsx");

            console.log(
              "Created file ",
              project_name + ".xlsx! Find attachments in the extract folder."
            );

            console.log('This file required "', apiCallCount, '" API calls.');
          })
          .catch(err => {
            console.error("Unable to create report.");
            console.error("Error: ", err);
          });
      })
    )
    .catch(err => {
      console.error("Error will retrieving project details: ", err);
    });
}

async function downloadAttachment(uri, folder, filename) {
  filename = filename.replace(/\/|\./g, "_");
  try {
    const response = await axiosInstance.get(uri, {
      responseType: "stream",
      "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    });
    const contentType = response.headers["content-disposition"];
    const savedFile = contentType
      .split("filename=")[1]
      .replace(/"/g, "")
      .trim();

    const format = savedFile.substring(savedFile.lastIndexOf(".") + 1);

    const currentPath = path.resolve(
      __dirname,
      "extract/" + folder,
      filename + "." + format
    );
    const writer = fs.createWriteStream(currentPath);

    writer.on("open", async function(fd) {
      response.data.pipe(writer);

      return new Promise((resolve, reject) => {
        writer.on("finish", resolve);
        writer.on("error", reject);
      });
    });
  } catch (err) {
    console.error("Error while downloading attachment: ", err);
  }
}

async function loadExpenses(expenseId, commonExpense) {
  let currentExpense = [];
  try {
    const res = await axiosInstance.get("expenses/" + expenseId);

    res.data.expense.line_items.forEach(x => {
      if (res.data.expense.project_id === projectId) {
        let itemExpense = Object.assign({}, commonExpense);
        // itemExpense['Currency'] = res.data.expense.currency_code
        // itemExpense['Line Item Id'] = x.line_item_id
        itemExpense["Description"] = x.description;
        itemExpense["Category"] = x.account_name;
        itemExpense["Amount in INR"] =
          x.amount * res.data.expense.exchange_rate;
        if (res.data.expense.currency_code === "USD") {
          itemExpense["Exchange Rate"] = res.data.expense.exchange_rate;
          itemExpense["Amount in USD"] = x.amount;
        }

        currentExpense.push(itemExpense);
      }
    });
  } catch (err) {
    console.error("Error while loading expenses: ", err);
  }
  return currentExpense;
}

async function loadBills(billId, commonBill, vendorId) {
  let currentBill = [];
  try {
    let vendor = vendorMap.get(vendorId);
    if (!vendor) {
      vendor = await axiosInstance.get("contacts/" + vendorId);
      vendorMap.set(vendorId, vendor);
    }

    // commonBill.Vendor = vendor.data.contact.contact_name

    const res = await axiosInstance.get("bills/" + billId);

    res.data.bill.line_items.forEach(x => {
      if (x.project_id === projectId) {
        let itemBill = Object.assign({}, commonBill);
        // itemBill['Currency'] = res.data.bill.currency_code

        // itemBill['Line Item Id'] = x.line_item_id
        itemBill["Description"] = vendor.data.contact.contact_name;
        itemBill["Category"] = x.account_name;
        itemBill["Amount in INR"] = x.item_total * res.data.bill.exchange_rate;
        if (res.data.bill.currency_code === "USD") {
          itemBill["Exchange Rate"] = res.data.bill.exchange_rate;
          itemBill["Amount in USD"] = x.item_total;
        }
        currentBill.push(itemBill);
      }
    });
  } catch (err) {
    console.error("Error while loading bills: ", err);
  }
  return currentBill;
}

async function fetchProjects(pageId) {
  try {
    let res = await axiosInstance.get("projects", {
      params: {
        organization_id: orgId,
        page: pageId,
        per_page: 25,
      },
    });
    projects = res.data.projects;
    if (res.data.page_context.has_more_page) {
      return projects.concat(await fetchProjects(pageId + 1));
    } else {
      return projects;
    }
  } catch (err) {
    console.error("Unable to retrieve list of available projects either.");
    console.error(err);
    return [];
  }
}

async function fetchExpenses(pageId) {
  try {
    let res = await axiosInstance.get("expenses", {
      params: {
        organization_id: orgId,
        page: pageId,
        per_page: 25,
      },
    });
    expenses = res.data.expenses;
    console.log("Fetched ", expenses.length, " expenses");
    if (res.data.page_context.has_more_page) {
      return expenses.concat(await fetchExpenses(pageId + 1));
    } else {
      return expenses;
    }
  } catch (err) {
    console.error("Unable to retrieve list of available expenses.");
    console.error(err);
    return [];
  }
}

async function fetchBills(pageId) {
  try {
    let res = await axiosInstance.get("bills", {
      params: {
        organization_id: orgId,
        page: pageId,
        per_page: 25,
      },
    });
    bills = res.data.bills;
    console.log("Fetched ", bills.length, " bills");
    if (res.data.page_context.has_more_page) {
      return bills.concat(await fetchBills(pageId + 1));
    } else {
      return bills;
    }
  } catch (err) {
    console.error("Unable to retrieve list of available bills.");
    console.error(err);
    return [];
  }
}

function buildProjectHeader(project) {
  return [
    ["Project Title:", project.description],
    ["ICFJ Internal Program ID:", project.custom_field_hash.cf_clientprojectid],
    ["Organization Name:", "Protograph Studio Pvt Ltd"],
    [
      "Dates Included in Reporting Period:",
      project.custom_field_hash.cf_from +
        " - " +
        project.custom_field_hash.cf_to,
    ],
    [""],
  ];
}

function buildProjectFooter() {
  return [
    [],
    [
      "Certification: By signing this report, I certify that it is true, complete, and accurate to the best of my knowledge. I am aware that any false, fictitious, or fraudulent information may subject me to criminal, civil, or administrative penalities.",
    ],
    ["Signature of Authorized Certifying Official:"],
    ["Name:"],
    ["Title:"],
    ["Date:"],
    [],
    [],
    ["INTERNAL USE ONLY"],
    [],
    ["Signature of ICFJ Program Director:"],
    ["Name:"],
    ["Date:"],
  ];
}

function createAdvanceRequestForm() {
  let ws = xlsx.utils.aoa_to_sheet(projectHeader);

  xlsx.utils.sheet_add_aoa(
    ws,
    [
      ["Budget Categories", "", "Estimated Expenses"],
      ["A. Subcontracts"],
      [
        "I. Total Direct Charges (Sum A - H)",
        "",
        { v: "", t: "n", z: dollarFormat },
      ],
      ["J. Indirect Charges", "", { v: "", t: "n", z: dollarFormat }],
      [
        "K. TOTAL PROJECT COSTS (I + J)",
        "",
        {
          f: "C9+C8",
          t: "n",
          z: dollarFormat,
        },
      ],
      [""],
      ["", "Period"],
      ["Total Award Amount"],
      [""],
      ["(a.) Cash Received", "", { v: "", t: "n", z: dollarFormat }],
      ["(b.) Cash Spent", "", { v: "", t: "n", z: dollarFormat }],
      [
        "(a. minus b.) Cash on Hand",
        "",
        { f: "C15-C16", t: "n", z: dollarFormat },
      ],
      [],
      ["Estimated Expenses", "", { f: "C10", t: "n", z: dollarFormat }],
      ["REQUESTED FUNDS", "", { f: "C19-C17", t: "n", z: dollarFormat }],
    ],
    { origin: -1 }
  );

  xlsx.utils.sheet_add_aoa(ws, projectFooter, { origin: -1 });

  return ws;
}

function createFinancialReport() {
  let ws = xlsx.utils.aoa_to_sheet(projectHeader);

  xlsx.utils.sheet_add_aoa(
    ws,
    [
      ["I", "II", "III", "IV", "V", "III + IV + V", "II - (III + IV + V)"],
      [
        "Budget Categories",
        "Approved Budget",
        "Actuals: Accumulated Through Prior Periods",
        "Actuals: This Period",
        "Adjustments",
        "Actuals: Accumulated Through This Period",
        "Remaining Budget",
      ],
      [
        "A. Subcontracts",
        "",
        "",
        "",
        { v: "0", t: "n", z: dollarFormat },
        { f: "SUM(C8:E8)", t: "n", z: dollarFormat },
        { f: "SUM(B8,-F8)", t: "n", z: dollarFormat },
      ],
      [
        "I. Total Direct Charges (Sum A - H)",
        { f: "SUM(B8)", t: "n", z: dollarFormat },
        { f: "SUM(C8)", t: "n", z: dollarFormat },
        { f: "SUM(D8)", t: "n", z: dollarFormat },
        { f: "SUM(E8)", t: "n", z: dollarFormat },
        { f: "SUM(F8)", t: "n", z: dollarFormat },
        { f: "SUM(G8)", t: "n", z: dollarFormat },
      ],
      [
        "J. Indirect Charges",
        "",
        "",
        "",
        "",
        { f: "SUM(C10,D10)", t: "n", z: dollarFormat },
        { f: "SUM(B10,-F10)", t: "n", z: dollarFormat },
      ],
      [
        "K. TOTAL PROJECT COSTS (I + J)",
        { f: "SUM(B9,B10)", t: "n", z: dollarFormat },
        { f: "SUM(C9,C10)", t: "n", z: dollarFormat },
        { f: "SUM(D9,D10)", t: "n", z: dollarFormat },
        { f: "SUM(E9,E10)", t: "n", z: dollarFormat },
        { f: "SUM(F9,F10)", t: "n", z: dollarFormat },
        { f: "SUM(G9,G10)", t: "n", z: dollarFormat },
      ],
    ],
    { origin: -1 }
  );

  xlsx.utils.sheet_add_aoa(ws, projectFooter, { origin: -1 });

  return ws;
}
