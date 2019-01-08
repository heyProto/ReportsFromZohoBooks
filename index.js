const axios = require('axios')
const fs = require('fs')
const path = require('path')
const xlsx = require('xlsx')

var args = process.argv.slice(2)

if (args.length != 4) {
  console.error('Incorrect number of arguments. Exiting process.')
  process.exit(1)
}

const authToken = args[0]
const orgId = args[1]
const projectId = args[2]
const allowAttachment = args[3]
let isDownloadable = false
let project_name = ''

if (allowAttachment) {
  switch (allowAttachment.toLowerCase()) {
    case 'y':
      isDownloadable = true
      break
    case 'n':
      isDownloadable = false
      break
    default:
      console.error('Invalid parameter. Expected (y/n)')
      process.exit(1)
  }
} else {
  console.error('Invalid parameter. Expected (y/n)')
  process.exit(1)
}

let axiosInstance = axios.create({
  baseURL: 'https://books.zoho.in/api/v3/',
  headers: {
    Authorization: 'Zoho-authtoken ' + authToken
  },
  timeout: 300000
})

axiosInstance
  .get('projects/' + projectId, {
    organization_id: orgId
  })
  .then(res => {
    console.log('Found project')
    findDetails()
  })
  .catch(err => {
    console.error('Unable to find project ID: ', projectId)
    console.error('Error while retrieving project: ', err)
    fetchProjects(1).then(projects => {
      console.error('Did you mean any of the following? ')
      projects.forEach(project => {
        console.error(project.project_id, ': ', project.project_name)
      })

      process.exit(1)
    })
  })

function findDetails () {
  axios
    .all([fetchExpenses(1), fetchBills(1) /*, fetchInvoices(1) */])
    .then(
      axios.spread((expenses, bills /*, invoices */) => {
        let expenseData = expenses
        let mappedExpense = expenseData.map(element => {
          let expenseId = element.expense_id
          let commonExpense = {
            'Transaction Date': element.date,
            // Type: 'Expense',
            // 'Reference Number': expenseId,
            Notes: element.reference_number
            // Status: element.status,
            // Vendor: '-'
          }
          if (element.has_attachment) {
            commonExpense['Attachment'] =
              'https://books.zoho.in/api/v3/expenses/' +
              expenseId +
              '/attachment'
            if (isDownloadable) {
              downloadAttachment(
                'expenses/' + expenseId + '/attachment',
                'expenses',
                commonExpense['Attachment']
              )
            }
          } else {
            commonExpense['Attachment'] = 'Not available'
          }

          return loadExpenses(expenseId, commonExpense)
        })
        let billData = bills
        let mappedBills = billData.map(element => {
          let billId = element.bill_id
          let commonBill = {
            'Transaction Date': element.date,
            // Type: 'Bill',
            // 'Reference Number': billId,
            Notes: element.bill_number
            // Status: element.status
          }

          if (element.has_attachment) {
            commonBill['Attachment'] =
              'https://books.zoho.in/api/v3/bills/' + billId + '/attachment'
            if (isDownloadable) {
              downloadAttachment(
                'bills/' + billId + '/attachment',
                'bills',
                commonBill['Attachment']
              )
            }
          } else {
            commonBill['Attachment'] = 'Not available'
          }

          return loadBills(billId, commonBill, element.vendor_id)
        })
        /*
        let invoiceData = invoices
        let mappedInvoices = invoiceData.map(element => {
          let invoiceId = element.invoice_id
          let commonInvoice = {
            'Transaction Date': element.date,
            Type: 'Invoice',
            'Reference Number': invoiceId,
            Id: element.invoice_number,
            Status: element.status
          }

          if (element.has_attachment) {
            commonInvoice['Attachment'] =
              'https://books.zoho.in/api/v3/invoices/' +
              invoiceId +
              '/attachment'
            if (isDownloadable) {
              downloadAttachment(
                'invoices/' + invoiceId + '/attachment',
                'invoices',
                commonInvoice['Attachment']
              )
            }
          } else {
            commonInvoice['Attachment'] = 'Not available'
          }

          return loadInvoices(invoiceId, commonInvoice)
        })
        */

        Promise.all([...mappedExpense, ...mappedBills /*, ...mappedInvoices */])
          .then(x => {
            let y = x.reduce((a, b) => a.concat(b), [])
            let subtotal_in_inr = 'C' + (y.length + 2)
            let subtotal_in_usd = 'E' + (y.length + 2)
            let fee_in_inr = 'C' + (y.length + 3)
            let fee_in_usd = 'E' + (y.length + 3)

            y = y.map((entry, i) => {
              let inr_cell = 'C' + (i + 2)
              let rate_cell = 'D' + (i + 2)
              let usd_amount = entry['Amount in USD']
              entry['Amount in USD'] = {
                f:
                  'IF(ISBLANK(' +
                  rate_cell +
                  '), ' +
                  usd_amount +
                  ', ' +
                  inr_cell +
                  ' / ' +
                  rate_cell +
                  ')',
                t: 'n'
              }
              return {
                'Transaction Date': entry['Transaction Date'],
                'Description of Good/Service': entry['Description'],
                'Amount in INR': entry['Amount in INR'],
                'Exchange Rate': entry['Exchange Rate'],
                'Amount in USD': entry['Amount in USD'],
                Notes: entry['Notes']
              }
            })
            let ws1 = xlsx.utils.json_to_sheet(y)
            xlsx.utils.sheet_add_aoa(
              ws1,
              [
                [
                  '',
                  'Subtotal',
                  { f: 'SUM(C2:C' + (y.length + 1) + ')', t: 'n' },
                  '',
                  { f: 'SUM(E2:E' + (y.length + 1) + ')', t: 'n' }
                ],
                [
                  '',
                  'Fee',
                  { f: '0.1*' + subtotal_in_inr, t: 'n' },
                  '',
                  { f: '0.1*' + subtotal_in_usd, t: 'n' }
                ],
                [
                  '',
                  'Total',
                  { f: subtotal_in_inr + '+' + fee_in_inr, t: 'n' },
                  '',
                  { f: subtotal_in_usd + '+' + fee_in_usd, t: 'n' }
                ]
              ],
              { origin: -1 }
            )

            let wb = xlsx.utils.book_new()
            xlsx.utils.book_append_sheet(wb, ws1, 'Expense Report')

            xlsx.writeFile(wb, project_name + '.xlsx')

            console.log(
              'Created file ',
              project_name + '.xlsx! Find attachments in the extract folder.'
            )
          })
          .catch(err => {
            console.error('Unable to create report.')
            console.error('Error: ', err)
          })
      })
    )
    .catch(err => {
      console.error('Error will retrieving project details: ', err)
    })
}

async function downloadAttachment (uri, folder, filename) {
  filename = filename.replace(/\/|\./g, '_')
  try {
    const response = await axiosInstance.get(uri, {
      responseType: 'stream',
      'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8'
    })
    const contentType = response.headers['content-disposition']
    const savedFile = contentType
      .split('filename=')[1]
      .replace(/"/g, '')
      .trim()

    const format = savedFile.substring(savedFile.lastIndexOf('.') + 1)

    const currentPath = path.resolve(
      __dirname,
      'extract/' + folder,
      filename + '.' + format
    )
    const writer = fs.createWriteStream(currentPath)

    writer.on('open', async function (fd) {
      response.data.pipe(writer)

      return new Promise((resolve, reject) => {
        writer.on('finish', resolve)
        writer.on('error', reject)
      })
    })
  } catch (err) {
    console.error('Error while downloading attachment: ', err)
  }
}

async function loadExpenses (expenseId, commonExpense) {
  let currentExpense = []
  try {
    const res = await axiosInstance.get('expenses/' + expenseId)
    if (project_name !== res.data.expense.project_name) {
      project_name = res.data.expense.project_name
    }

    res.data.expense.line_items.forEach(x => {
      if (res.data.expense.project_id === projectId) {
        let itemExpense = Object.assign({}, commonExpense)
        // itemExpense['Currency'] = res.data.expense.currency_code
        // itemExpense['Line Item Id'] = x.line_item_id
        itemExpense['Description'] = x.description + ' - ' + x.account_name
        // itemExpense['Account Name'] = x.account_name
        itemExpense['Amount in INR'] = x.amount * res.data.expense.exchange_rate
        if (res.data.expense.currency_code === 'USD') {
          itemExpense['Exchange Rate'] = res.data.expense.exchange_rate
          itemExpense['Amount in USD'] = x.amount
        }

        currentExpense.push(itemExpense)
      }
    })
  } catch (err) {
    console.error('Error while loading expenses: ', err)
  }
  return currentExpense
}

async function loadBills (billId, commonBill, vendorId) {
  let currentBill = []
  try {
    const vendor = await axiosInstance.get('contacts/' + vendorId)

    // commonBill.Vendor = vendor.data.contact.contact_name

    const res = await axiosInstance.get('bills/' + billId)

    res.data.bill.line_items.forEach(x => {
      if (x.project_id === projectId) {
        if (project_name !== x.project_name) {
          project_name = x.project_name
        }
        let itemBill = Object.assign({}, commonBill)
        // itemBill['Currency'] = res.data.bill.currency_code

        // itemBill['Line Item Id'] = x.line_item_id
        itemBill['Description'] =
          vendor.data.contact.contact_name + ' - ' + x.account_name
        // itemBill['Account Name'] = x.account_name
        itemBill['Amount in INR'] = x.item_total * res.data.bill.exchange_rate
        if (res.data.bill.currency_code === 'USD') {
          itemBill['Exchange Rate'] = res.data.bill.exchange_rate
          itemBill['Amount in USD'] = x.item_total
        }
        currentBill.push(itemBill)
      }
    })
  } catch (err) {
    console.error('Error while loading bills: ', err)
  }
  return currentBill
}

async function loadInvoices (invoiceId, commonInvoice) {
  let currentInvoice = []
  try {
    const res = await axiosInstance.get('invoices/' + invoiceId)

    res.data.invoice.line_items.forEach(x => {
      if (x.project_id === projectId) {
        let itemInvoice = Object.assign({}, commonInvoice)
        itemInvoice['Currency'] = res.data.invoice.currency_code
        itemInvoice['Currency Rate'] = res.data.invoice.exchange_rate
        itemInvoice['Line Item Id'] = x.line_item_id
        itemInvoice['Description'] = x.name || x.description
        itemInvoice['Account Name'] = x.account_name
        itemInvoice['Amount in INR'] =
          x.item_total * res.data.invoice.exchange_rate
        currentInvoice.push(itemInvoice)
      }
    })
  } catch (err) {
    console.error('Error while loading invoices: ', err)
  }
  return currentInvoice
}

async function fetchProjects (pageId) {
  try {
    let res = await axiosInstance.get('projects', {
      params: {
        organization_id: orgId,
        page: pageId,
        per_page: 25
      }
    })
    projects = res.data.projects
    if (res.data.page_context.has_more_page) {
      return projects.concat(await fetchProjects(pageId + 1))
    } else {
      return projects
    }
  } catch (err) {
    console.error('Unable to retrieve list of available projects either.')
    console.error(err)
    return []
  }
}

async function fetchExpenses (pageId) {
  try {
    let res = await axiosInstance.get('expenses', {
      params: {
        organization_id: orgId,
        page: pageId,
        per_page: 25
      }
    })
    expenses = res.data.expenses
    console.log('Fetched ', expenses.length, ' expenses')
    if (res.data.page_context.has_more_page) {
      return expenses.concat(await fetchExpenses(pageId + 1))
    } else {
      return expenses
    }
  } catch (err) {
    console.error('Unable to retrieve list of available expenses.')
    console.error(err)
    return []
  }
}

async function fetchBills (pageId) {
  try {
    let res = await axiosInstance.get('bills', {
      params: {
        organization_id: orgId,
        page: pageId,
        per_page: 25
      }
    })
    bills = res.data.bills
    console.log('Fetched ', bills.length, ' bills')
    if (res.data.page_context.has_more_page) {
      return bills.concat(await fetchBills(pageId + 1))
    } else {
      return bills
    }
  } catch (err) {
    console.error('Unable to retrieve list of available bills.')
    console.error(err)
    return []
  }
}

async function fetchInvoices (pageId) {
  try {
    let res = await axiosInstance.get('invoices', {
      params: {
        organization_id: orgId,
        page: pageId,
        per_page: 25
      }
    })
    invoices = res.data.invoices
    console.log('Fetched ', invoices.length, ' invoices')
    if (res.data.page_context.has_more_page) {
      return invoices.concat(await fetchInvoices(pageId + 1))
    } else {
      return invoices
    }
  } catch (err) {
    console.error('Unable to retrieve list of available invoices.')
    console.error(err)
    return []
  }
}
