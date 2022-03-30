import { config } from 'dotenv-flow'
config()

import Airtable from 'airtable'

const airtable = new Airtable({ apiKey: process.env.AIRTABLE_API_KEY }).base(process.env.AIRTABLE_BASE as string)
