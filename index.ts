import express from "express"
import fileUpload from "express-fileupload"
import {read} from 'xlsx'
import {z, ZodType} from "zod";

const app = express()
app.use(express.json())
app.use(fileUpload())

type File = Record<string, string>
type FileEntry = [string, string | number]

const mandatoryFields = new Set([
    "Customer", "Cust No'", "Project Type",
    "Quantity", "Price Per Item", "Item Price Currency",
    "Total Price", "Invoice Currency", "Status"
])

const invoiceDataFormat: { [key: string]: ZodType } = {
    "Customer": z.string(),
    "Cust No'": z.number().int().gte(10000).lte(99999), // 5-digit integer
    "Project Type": z.string().max(20),
    "Quantity": z.number(),
    "Price Per Item": z.number(),
    "Item Price Currency": z.string().toUpperCase().length(3), // 3 uppercase letter identifier
    "Total Price": z.number(),
    "Invoice Currency": z.string().toUpperCase().length(3),
    "Status": z.enum(["Ready", "Done"]), // status can be ready or done
}

const columnLetters = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
    'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
    'U', 'V', 'W', 'X', 'Y', 'Z'
];

const CHAR_OFFSET = 65

function getRightCellIndex(cell: string) {
    // Cells, like "A1", "B2" etc.
    const columnIndex = cell.charCodeAt(0) - CHAR_OFFSET // converting uppercase letter to array index
    if (columnIndex >= 0 && columnIndex < columnLetters.length - 1) {
        return "".concat(columnLetters[columnIndex + 1]!, cell[1]!)
    } else throw new Error("Wrong column index")
}

function parseInvoicingDate(dateString: string) {
    const monthsArray: {[key: string]: string} = {
        Jan: '01',
        Feb: '02',
        Mar: '03',
        Apr: '04',
        May: '05',
        Jun: '06',
        Jul: '07',
        Aug: '08',
        Sep: '09',
        Oct: '10',
        Nov: '11',
        Dec: '12',
    }

    // Assuming format like: "Sep 2023"
    const dateArray = dateString.split(" ")
    const [ month, year] = [ dateArray[0]!, dateArray[1]! ]
    return `${year}-${monthsArray[month]}`
}

function parseCurrencyRates(file: File) {
    const currencyRates: { [key: string]: string } = {}

    // Selecting string cells, that end with word "Rate"
    const rateCells = fileToEntries(file).filter((value: FileEntry) =>
        typeof value[1] === "string" ? value[1].endsWith("Rate") : false
    )

    rateCells.forEach((value) => {
        const [cellIndex, cellValue] = [value[0], value[1]]
        const rateName = cellValue.split(" ")[0]! // first word before "Rate"
        const rightCellIndex = getRightCellIndex(cellIndex) // assuming rate value is right next to its name
        currencyRates[rateName] = file[rightCellIndex] || "No value"
    })

    return currencyRates
}

function parseInvoicesData(file: File) {
    const fileEntries: FileEntry[] = fileToEntries(file)
    const invoicesData: {}[] = []

    // file should be validated before this
    const statusEntry =
        fileEntries.find(element => element[1] === "Status") as FileEntry
    const invoiceNoEntry =
        fileEntries.find(element => element[1] === "Invoice #") as FileEntry

    const statusCell = statusEntry[0]
    const [ statusCol, statusRow ] = rowColFromString(statusCell)

    const invoiceNoCell = invoiceNoEntry[0]
    const [ invoiceNoCol , invoiceNoRow ] = rowColFromString(invoiceNoCell)

    if (statusRow !== invoiceNoRow) {
        throw new Error("Broken table")
    }

    let headerRowIndex: number = statusRow

    let columns = new Map()

    // Retrieving all column names
    for (let i = 0; i < columnLetters.length; i++) {
        const cellIndex = `${columnLetters[i]}${headerRowIndex}`
        if (file[cellIndex]) {
            columns.set(cellIndex, file[cellIndex])
        }
    }

    const entriesToProcess = fileEntries.filter(value => {
        const [ entryIndex, entryValue] = value
        const [ valueCol, valueRow ] = rowColFromString(entryIndex)

        if (valueRow > statusRow) {
            if (valueCol === statusCol && entryValue === "Ready") {
                return true
            } else {
                return valueCol === invoiceNoCol && // invoiceNo is a string of format "INV00000000"
                typeof entryValue === "string" &&
                entryValue.startsWith("INV") &&
                entryValue.length === 11;
            }
        } else {
            return false
        }
    })

    const rowsToProcess: number[] = entriesToProcess.map(element => parseInt(element[0].substring(1)))

    rowsToProcess.forEach(row => {
        const entry: { [key: string]: string | number | string[] } = {
            validationErrors: []
        }

        for (let j = 0; j < columnLetters.length; j++) {
            const cellIndex = `${columnLetters[j]}${row}`
            const columnNameIndex = `${columnLetters[j]}${headerRowIndex}`
            const columnName = columns.get(columnNameIndex)

            const cellValue = file[cellIndex]!

            // if this column is mandatory, we should validate it
            if (mandatoryFields.has(columnName)) {
                const result = invoiceDataFormat[columnName]!.safeParse(cellValue)

                if (!result.success) {
                    const validationErrorsArray =
                        result.error.errors.map(errorObject => `${columnName}: ${errorObject.message}`)

                    entry.validationErrors = [...(entry.validationErrors as string[] || []), ...validationErrorsArray]
                }
            }

            entry[columnName] = cellValue
        }

        invoicesData.push(entry)
    })

    return invoicesData
}

function validateSheetStructure(file: File) {
    const fileEntries: FileEntry[] = fileToEntries(file)
    const statusEntry =
        fileEntries.find(element => element[1] === "Status")
    const invoiceNoEntry =
        fileEntries.find(element => element[1] === "Invoice #")

    if (!statusEntry) {
        throw new Error("No 'Status' cell in sheet")
    }

    if (!invoiceNoEntry) {
        throw new Error("No 'Invoice #' cell in sheet")
    }

    if (typeof file[`A1`] !== "string") {
        throw new Error("Wrong Invoice Date")
    }

    mandatoryFields.forEach(mandatoryField => {
        if (fileEntries.find(cell => cell[1] === mandatoryField) === undefined) {
            throw new Error("Mandatory field is missing")
        }
    })
}

function calculateTotalInvoiceCurrency(
    invoicesData: { [key: string]: string }[],
    currencyRates: { [key: string]: string })
{
    return invoicesData.map((invoiceData: { [key: string]: string }) => {
        const invoiceCurrency = invoiceData["Invoice Currency"]!

        if (currencyRates[invoiceCurrency] !== undefined) {
            invoiceData["Invoice Total"] = `${parseInt(invoiceData["Total Price"]!) * parseFloat(currencyRates[invoiceCurrency]!)}`
        } else {
            invoiceData["Invoice Total"] = "No such currency defined"
        }

        return invoiceData
    })
}

function fileToEntries(file: File) {
    return Object.entries(file)
}

function rowColFromString(cell: string): [string, number] {
    // First letter in cell name is a column
    return [cell[0]!, parseInt(cell.substring(1))]
}

app.post("/", (req: any, res) => {
    // Parsing file to object, which contains only cells with defined values
    const file: File = {}
    try {
        Object.entries(read(req.files.data.data).Sheets?.Sheet1!)
            .filter(value => value[1].v !== undefined) // filtering out meta fields
            .forEach(element => file[element[0]] = element[1].v) // we need only cell values
    } catch (e) {
        throw new Error("Error parsing file: " + e)
    }

    validateSheetStructure(file)

    // Assuming that invoicing month can be only in this exact cell and only in provided format
    const invoicingMonth = parseInvoicingDate(file[`A1`]!)

    const invoicingMonthParameter = req.body.invoicingMonth
    if (invoicingMonth !== invoicingMonthParameter) {
        throw new Error("Invoicing month param dont match file invoicing month")
    }

    const currencyRates = parseCurrencyRates(file)

    let invoicesData = parseInvoicesData(file)

    invoicesData = calculateTotalInvoiceCurrency(invoicesData, currencyRates)

    const outputObj = {
        InvoicingMonth: invoicingMonth,
        currencyRates: currencyRates,
        invoicesData: invoicesData
    }

    res.send(outputObj)
})

app.listen(3000, ()=> {
    console.log("server started")
})

