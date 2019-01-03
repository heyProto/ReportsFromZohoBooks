const axios = require("axios");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

var args = process.argv.slice(2);

const authToken = args[0];
const orgId = args[1];

let axiosInstance = axios.create({
  baseURL: "https://books.zoho.in/api/v3/",
  headers: {
    Authorization: "Zoho-authtoken " + authToken,
  },
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
        console.log("HERE");
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

      Promise.all([...mappedExpense, ...mappedBills, ...mappedInvoices]).then(
        x => {
          let y = x.reduce((a, b) => a.concat(b), []);

          let ws1 = xlsx.utils.json_to_sheet(y);
          //   let ws3 = xlsx.utils.json_to_sheet(invoiceData);

          let wb = xlsx.utils.book_new();
          xlsx.utils.book_append_sheet(wb, ws1, "Expenses");
          //   xlsx.utils.book_append_sheet(wb, ws3, "Invoices");

          xlsx.writeFile(wb, "zoho.xlsx");
        }
      );
    })
  );

async function downloadAttachment(uri, folder, filename) {
  console.log(uri);
  const response = await axiosInstance.get(uri, {
    responseType: "stream",
    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
  });

  const contentType = response.headers["content-type"];
  const format = contentType.substring(contentType.lastIndexOf("/") + 1);
  const currentPath = path.resolve(
    __dirname,
    "extract/" + folder,
    filename + "." + format
  );
  console.log(currentPath);
  const writer = fs.createWriteStream(currentPath);

  writer.on("open", async function(fd) {
    response.data.pipe(writer);

    return new Promise((resolve, reject) => {
      writer.on("finish", resolve);
      writer.on("error", reject);
    });
  });
}

async function loadExpenses(expenseId, commonExpense) {
  let currentExpense = [];
  const res = await axiosInstance.get("expenses/" + expenseId);

  res.data.expense.line_items.forEach(x => {
    let itemExpense = Object.assign({}, commonExpense);
    itemExpense["Currency"] = res.data.expense.currency_code;
    itemExpense["Currency Rate"] = res.data.expense.exchange_rate;
    itemExpense["Line Item Id"] = x.line_item_id;
    itemExpense["Description"] = x.description;
    itemExpense["Account Name"] = x.account_name;
    itemExpense["Amount in INR"] = x.amount * res.data.expense.exchange_rate;
    currentExpense.push(itemExpense);
  });
  return currentExpense;
}

async function loadBills(billId, commonBill, vendorId) {
  let currentBill = [];

  const vendor = await axiosInstance.get("contacts/" + vendorId);

  commonBill.Vendor = vendor.data.contact.contact_name;

  const res = await axiosInstance.get("bills/" + billId);

  res.data.bill.line_items.forEach(x => {
    let itemBill = Object.assign({}, commonBill);
    itemBill["Currency"] = res.data.bill.currency_code;
    itemBill["Currency Rate"] = res.data.bill.exchange_rate;
    itemBill["Line Item Id"] = x.line_item_id;
    itemBill["Description"] = x.description;
    itemBill["Account Name"] = x.account_name;
    itemBill["Amount in INR"] = x.item_total * res.data.bill.exchange_rate;
    currentBill.push(itemBill);
  });
  return currentBill;
}

async function loadInvoices(invoiceId, commonInvoice) {
  let currentInvoice = [];
  const res = await axiosInstance.get("invoices/" + invoiceId);

  currentInvoice = res.data.invoice.line_items.map(x => {
    let itemInvoice = Object.assign({}, commonInvoice);
    itemInvoice["Currency"] = res.data.invoice.currency_code;
    itemInvoice["Currency Rate"] = res.data.invoice.exchange_rate;
    itemInvoice["Line Item Id"] = x.line_item_id;
    itemInvoice["Description"] = x.name || x.description;
    itemInvoice["Account Name"] = x.account_name;
    itemInvoice["Amount in INR"] =
      x.item_total * res.data.invoice.exchange_rate;
    return itemInvoice;
  });
  return currentInvoice;
}
