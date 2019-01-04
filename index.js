const axios = require("axios");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

var args = process.argv.slice(2);

if (args.length != 3) {
  console.error("Incorrect number of arguments. Exiting process.");
  process.exit(1);
}

const authToken = args[0];
const orgId = args[1];
const projectId = args[2];

let axiosInstance = axios.create({
  baseURL: "https://books.zoho.in/api/v3/",
  headers: {
    Authorization: "Zoho-authtoken " + authToken,
  },
});

axiosInstance
  .get("projects/" + projectId, {
    organization_id: orgId,
  })
  .then(res => {
    console.log("Found project");
  })
  .catch(err => {
    console.error("Unable to find project ID: ", projectId);
    console.error("Error while retrieving project: ", err);
    axiosInstance
      .get("projects", {
        organization_id: orgId,
      })
      .then(res => {
        console.error("Did you mean any of the following? ");
        res.data.projects.forEach(project => {
          console.error(project.project_id, ": ", project.project_name);
        });

        process.exit(1);
      })
      .catch(err => {
        console.error("Unable to retrieve list of  available projects either.");
        console.error(err);
      });
  });

axios
  .all([
    axiosInstance.get("expenses", {
      organization_id: orgId,
    }),
    axiosInstance.get("bills", {
      organization_id: orgId,
    }),
    axiosInstance.get("invoices", {
      organization_id: orgId,
    }),
  ])
  .then(
    axios.spread((expenses, bills, invoices) => {
      let expenseData = expenses.data.expenses;
      let mappedExpense = expenseData.map(element => {
        let expenseId = element.expense_id;
        let commonExpense = {
          "Transaction Date": element.date,
          Type: "Expense",
          "Reference Number": expenseId,
          Id: element.reference_number,
          Status: element.status,
          Vendor: "-",
        };
        if (element.has_attachment) {
          commonExpense["Attachment"] = element.reference_number;
          downloadAttachment(
            "expenses/" + expenseId + "/attachment",
            "expenses",
            commonExpense["Attachment"]
          );
        } else {
          commonExpense["Attachment"] = "Not available";
        }

        return loadExpenses(expenseId, commonExpense);
      });
      let billData = bills.data.bills;
      let mappedBills = billData.map(element => {
        let billId = element.bill_id;
        let commonBill = {
          "Transaction Date": element.date,
          Type: "Bill",
          "Reference Number": billId,
          Id: element.bill_number,
          Status: element.status,
        };

        if (element.has_attachment) {
          commonBill["Attachment"] = element.bill_number;
          downloadAttachment(
            "bills/" + billId + "/attachment",
            "bills",
            commonBill["Attachment"]
          );
        } else {
          commonBill["Attachment"] = "Not available";
        }

        return loadBills(billId, commonBill, element.vendor_id);
      });
      let invoiceData = invoices.data.invoices;
      let mappedInvoices = invoiceData.map(element => {
        let invoiceId = element.invoice_id;
        let commonInvoice = {
          "Transaction Date": element.date,
          Type: "Invoice",
          "Reference Number": invoiceId,
          Id: element.invoice_number,
          Status: element.status,
        };

        if (element.has_attachment) {
          commonInvoice["Attachment"] = element.invoice_number;
          downloadAttachment(
            "invoices/" + invoiceId + "/attachment",
            "invoices",
            commonInvoice["Attachment"]
          );
        } else {
          commonInvoice["Attachment"] = "Not available";
        }

        return loadInvoices(invoiceId, commonInvoice);
      });

      Promise.all([...mappedExpense, ...mappedBills, ...mappedInvoices])
        .then(x => {
          let y = x.reduce((a, b) => a.concat(b), []);

          let ws1 = xlsx.utils.json_to_sheet(y);
          //   let ws3 = xlsx.utils.json_to_sheet(invoiceData);

          let wb = xlsx.utils.book_new();
          xlsx.utils.book_append_sheet(wb, ws1, "Expenses");
          //   xlsx.utils.book_append_sheet(wb, ws3, "Invoices");

          xlsx.writeFile(wb, "zoho.xlsx");

          console.log(
            "Created file zoho.xlsx! Find attachments in the extract folder."
          );
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

async function downloadAttachment(uri, folder, filename) {
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
        itemExpense["Currency"] = res.data.expense.currency_code;
        itemExpense["Currency Rate"] = res.data.expense.exchange_rate;
        itemExpense["Line Item Id"] = x.line_item_id;
        itemExpense["Description"] = x.description;
        itemExpense["Account Name"] = x.account_name;
        itemExpense["Amount in INR"] =
          x.amount * res.data.expense.exchange_rate;
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
    const vendor = await axiosInstance.get("contacts/" + vendorId);

    commonBill.Vendor = vendor.data.contact.contact_name;

    const res = await axiosInstance.get("bills/" + billId);

    res.data.bill.line_items.forEach(x => {
      if (x.project_id === projectId) {
        let itemBill = Object.assign({}, commonBill);
        itemBill["Currency"] = res.data.bill.currency_code;
        itemBill["Currency Rate"] = res.data.bill.exchange_rate;
        itemBill["Line Item Id"] = x.line_item_id;
        itemBill["Description"] = x.description;
        itemBill["Account Name"] = x.account_name;
        itemBill["Amount in INR"] = x.item_total * res.data.bill.exchange_rate;
        currentBill.push(itemBill);
      }
    });
  } catch (err) {
    console.error("Error while loading bills: ", err);
  }
  return currentBill;
}

async function loadInvoices(invoiceId, commonInvoice) {
  let currentInvoice = [];
  try {
    const res = await axiosInstance.get("invoices/" + invoiceId);

    res.data.invoice.line_items.forEach(x => {
      if (x.project_id === projectId) {
        let itemInvoice = Object.assign({}, commonInvoice);
        itemInvoice["Currency"] = res.data.invoice.currency_code;
        itemInvoice["Currency Rate"] = res.data.invoice.exchange_rate;
        itemInvoice["Line Item Id"] = x.line_item_id;
        itemInvoice["Description"] = x.name || x.description;
        itemInvoice["Account Name"] = x.account_name;
        itemInvoice["Amount in INR"] =
          x.item_total * res.data.invoice.exchange_rate;
        currentInvoice.push(itemInvoice);
      }
    });
  } catch (err) {
    console.error("Error while loading invoices: ", err);
  }
  return currentInvoice;
}
