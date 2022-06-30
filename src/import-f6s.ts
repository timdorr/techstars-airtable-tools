import { config } from 'dotenv-flow'
config()

import XLSX from 'xlsx'
import Airtable, { FieldSet } from 'airtable'
import { parse } from 'date-fns'

import { F6SPipeline, F6SCompany, PersonNumber } from './f6s'

const sheet = XLSX.readFile(process.argv.slice(-1)[0])
const apps = XLSX.utils.sheet_to_json<F6SPipeline>(sheet.Sheets['Pipeline'], { range: 1 })

const pipelines: Record<string, F6SCompany[]> = {}
for (const pipeline of sheet.SheetNames) {
  if (pipeline == 'Pipeline') continue

  pipelines[pipeline] = XLSX.utils.sheet_to_json<F6SCompany>(sheet.Sheets[pipeline], { range: 1 })
}

const airtable = new Airtable({ apiKey: process.env.AIRTABLE_API_KEY }).base(
  process.env.AIRTABLE_SOURCING_BASE as string
)

async function upsert(table: string, keyField: string, keyValue: string | number, fields: FieldSet): Promise<string> {
  try {
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
  } catch (e) {
    console.error(e)
    return e instanceof Error ? e.message : ''
  }
}

async function main() {
  for (const application of apps) {
    const pipeline =
      application['Pipeline'].length > 31
        ? `${application['Pipeline'].slice(0, 14)}...${application['Pipeline'].slice(-14)}`
        : application['Pipeline']

    const app = pipelines[pipeline].find(company => company['User ID'] == application['Startup ID'])
    if (!app) {
      console.log(`!!${application['Item Name']} not found!!`)
      continue
    }

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
      Program: application['Pipeline'],
      'Date Created': parse(
        `${app['Date Created']} ${app['Time Created']}`,
        'dd/MM/yyyy HH:mm:ss',
        new Date()
      ).toISOString(),
      Company: [company_id]
    })

    await Array.from({ length: 5 }, async (_, n) => {
      const num: PersonNumber = (n + 1) as PersonNumber

      if (!app[`Person ${num} User ID`]) return
      await upsert('Founders', 'F6S User ID', parseInt(app[`Person ${num} User ID`]), {
        Name: `${app[`Person ${num} Name`]} ${app[`Person ${num} Surname`]}`,
        Email: app[`Person ${num} Email`],
        Phone: app[`Person ${num} Phone`],
        Location: `${app[`Person ${num} City`]}${
          app[`Person ${num} Country`] && app[`Person ${num} Country`] != 'United States'
            ? `, ${app[`Person ${num} Country`]}`
            : ''
        }`,
        Role: app[`Person ${num} Role`],
        Skills: app[`Person ${num} Skills or Markets`].split(', ').filter(Boolean),
        Facebook: app[`Person ${num} Facebook`],
        Linkedin: app[`Person ${num} Linkedin`],
        Twitter: app[`Person ${num} Twitter`],
        Website: app[`Person ${num} Website`],
        Description: app[`Person ${num} Brief Description`],
        Experience: app[`Person ${num} Experience`],
        Companies: [company_id]
      })
    })
  }
}

main()
