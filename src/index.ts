import { config } from 'dotenv-flow'
config()

import XLSX from 'xlsx'
import Airtable, { FieldSet, Records } from 'airtable'
import { parse } from 'date-fns'

import { F6SCompany } from './f6s'

const sheet = XLSX.readFile(process.argv.slice(-1)[0])
const apps = XLSX.utils.sheet_to_json<F6SCompany>(sheet.Sheets[sheet.SheetNames[1]], { range: 1 })

const airtable = new Airtable({ apiKey: process.env.AIRTABLE_API_KEY }).base(process.env.AIRTABLE_BASE as string)

async function upsert(table: string, keyField: string, keyValue: string | number, fields: FieldSet): Promise<string> {
  const records = await airtable(table)
    .select({ filterByFormula: `{${keyField}} = '${keyValue}'` })
    .firstPage()

  if (records.length == 0) {
    const result = await airtable(table).create(
      {
        [keyField]: keyValue,
        ...fields
      },
      { typecast: true }
    )

    return result.id
  } else {
    const result = await airtable(table).update(records[0].id, fields, { typecast: true })
    return result.id
  }
}

async function main() {
  try {
    for (const app of apps) {
      console.log(`Processing application for ${app['Startup/Person Name']}`)

      const company_id = await upsert('Companies', 'F6S Company ID', parseInt(app['User ID']), {
        Name: app['Startup/Person Name'],
        Description: app['Brief Description'],
        'Product Video': app['Product Video'],
        'Team Video': app['Team Video'],
        Location: `${app.City}${app.Country && app.Country != 'United States' ? `, ${app.Country}` : ''}`,
        Website: app.Website,
        Facebook: app.Facebook,
        Twitter: app.Twitter,
        Linkedin: app.Linkedin,
        Incorporated: app['Are you registered or incorporated?'] == 'Yes',
        'Where Incorporated': app['Where are you registered or incorporated?'],
        'How Far Along': app['How far along are you?'],
        'Money Raised': parseInt(app['How much money raised since start?'].replace(/\D/g, '')),
        'Key Customers': app['Key customers/users?'],
        Raising: app.Raising == 'Yes',
        'Amount Raising': parseInt(app['Amount Raising']),
        Valuation: parseInt(app.Valuation),
        'Fund Stage': app['Fund Stage']
      })

      await upsert('Applications', 'F6S Application ID', parseInt(app['Application ID']), {
        'Primary Contact Name': app['Primary Contact Name'],
        'Primary Contact Title': app['Primary Contact Title'],
        'Primary Contact Email': app['Primary Contact Email Address'],
        Status: app.Status,
        'Complete %': parseInt(app['Complete %']) / 100,
        Program: sheet.SheetNames[1],
        'Date Created': parse(
          `${app['Date Created']} ${app['Time Created']}`,
          'dd/MM/yyyy HH:mm:ss',
          new Date()
        ).toISOString(),
        Company: [company_id]
      })
    }
  } catch (e) {
    console.error(e)
  }
}

main()
