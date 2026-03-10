#!/usr/bin/env node
// import_march.js — รัน via: node scripts/import_march.js
const XLSX = require('xlsx')
const { Pool } = require('pg')
const path = require('path')

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes('railway.internal') ? false : { rejectUnauthorized: false }
})

function parseDate(name) {
  const clean = name.replace(/\(.*\)/g, '').trim()
  const m = clean.match(/(\d+)\s*มี\.ค\./)
  if (!m) return null
  return '2026-03-' + String(m[1]).padStart(2, '0')
}

async function main() {
  const client = await pool.connect()
  console.log('✅ DB connected')

  const check = await client.query('SELECT COUNT(*) FROM delivery_rows')
  console.log('Current rows in DB:', check.rows[0].count)

  const filePath = path.join(__dirname, '..', 'import_data.xlsx')
  const wb = XLSX.readFile(filePath)
  let total = 0, skipped = 0

  for (const sheetName of wb.SheetNames) {
    const date = parseDate(sheetName)
    if (!date) { console.log(`SKIP (no date): ${sheetName}`); skipped++; continue }

    const ws = wb.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })
    const dataRows = rows.filter(r => typeof r[0] === 'number' && r[0] > 0)
    if (!dataRows.length) { console.log(`SKIP (empty): ${sheetName}`); skipped++; continue }

    console.log(`${sheetName} → ${date} : ${dataRows.length} rows`)

    for (const row of dataRows) {
      await client.query(
        `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker)
         VALUES ($1,$2,$3,$4,$5,$6,false)`,
        [date, 'inter', String(row[1]||'').trim(), String(row[2]||'').trim(), String(row[4]||'').trim(), parseInt(row[3])||0]
      )
      total++
    }
  }

  console.log(`\n✅ Done! inserted: ${total} rows | skipped: ${skipped} sheets`)
  client.release()
  await pool.end()
}

main().catch(e => { console.error('❌', e.message); process.exit(1) })
