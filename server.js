const express = require('express')
const multer = require('multer')
const path = require('path')

const XLSX = require('xlsx')
const { Pool } = require('pg')

const app = express()
const PORT = process.env.PORT || 3100

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } })

// OCR ผ่าน OCR Proxy บน Mac (cloudflare tunnel) — retry 3 ครั้ง
async function ocrViaGateway(base64, mimeType) {
  const PROXY_URL = process.env.OPENCLAW_GATEWAY_URL
  const SECRET = process.env.OPENCLAW_GATEWAY_TOKEN || 'newlife2026'
  if (!PROXY_URL) throw new Error('OPENCLAW_GATEWAY_URL not set')

  let lastErr
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const resp = await fetch(`${PROXY_URL}/ocr`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ base64, mimeType, secret: SECRET }),
        signal: AbortSignal.timeout(90000)
      })
      const data = await resp.json()
      if (!data.ok) throw new Error(data.error || 'OCR failed')
      return data.result
    } catch (e) {
      lastErr = e
      console.error(`OCR attempt ${attempt} failed:`, e.message)
      if (attempt < 3) await new Promise(r => setTimeout(r, 2000 * attempt))
    }
  }
  throw lastErr
}

// ── DB ──
const dbUrl = process.env.DATABASE_URL || ''
const needSsl = dbUrl && !dbUrl.includes('railway.internal') && !dbUrl.includes('localhost')
console.log(`DB URL host: ${dbUrl.split('@')[1]?.split('/')[0] || 'none'} | SSL: ${needSsl}`)
const pool = new Pool({
  connectionString: dbUrl,
  ssl: needSsl ? { rejectUnauthorized: false } : false
})

async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS delivery_rows (
      id SERIAL PRIMARY KEY,
      session_date DATE NOT NULL DEFAULT CURRENT_DATE,
      carrier VARCHAR(10) NOT NULL,
      shop_name TEXT DEFAULT '',
      province TEXT DEFAULT '',
      invoice_no TEXT DEFAULT '',
      quantity INTEGER DEFAULT 0,
      red_sticker BOOLEAN DEFAULT FALSE,
      sort_order INTEGER DEFAULT 0,
      session_slot VARCHAR(10) DEFAULT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    )
  `)
  // เพิ่มคอลัมน์ session_slot ถ้ายังไม่มี (migration)
  await pool.query(`
    ALTER TABLE delivery_rows ADD COLUMN IF NOT EXISTS session_slot VARCHAR(10) DEFAULT NULL
  `)
  console.log('✅ DB ready')
}

app.use(express.json())
app.use(express.static(path.join(__dirname, 'public')))

// /view → view.html
app.get('/view', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'view.html'))
})

// ── GET rows (by date) ──
app.get('/api/rows', async (req, res) => {
  try {
    const date = req.query.date || new Date().toISOString().slice(0, 10)
    const sort = req.query.sort || 'province' // 'created' = ลำดับที่ถ่าย, 'province' = เรียง province
    const orderBy = sort === 'created'
      ? 'carrier, session_slot NULLS LAST, created_at, id'
      : 'carrier, session_slot NULLS LAST, province, shop_name, id'
    const r = await pool.query(
      `SELECT * FROM delivery_rows WHERE session_date = $1 ORDER BY ${orderBy}`,
      [date]
    )
    res.json({ ok: true, rows: r.rows, date })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── Admin: query rows by created_at date (for recovery) ──
app.get('/api/admin/recover', async (req, res) => {
  const { secret, created_date } = req.query
  if (secret !== 'newlife2026') return res.status(403).json({ ok: false })
  const r = await pool.query(
    `SELECT id, shop_name, province, invoice_no, quantity, red_sticker, session_date, session_slot, created_at
     FROM delivery_rows WHERE DATE(created_at) = $1 ORDER BY id`,
    [created_date]
  )
  res.json({ ok: true, rows: r.rows })
})

// ── Search rows across all dates ──
app.get('/api/rows/search', async (req, res) => {
  try {
    const { q, carrier } = req.query
    if (!q || q.trim().length < 1) return res.json({ ok: true, rows: [] })
    const like = '%' + q.trim() + '%'
    const r = await pool.query(
      `SELECT * FROM delivery_rows
       WHERE (shop_name ILIKE $1 OR province ILIKE $1 OR invoice_no ILIKE $1)
       ${carrier ? 'AND carrier = $2' : ''}
       ORDER BY session_date DESC, carrier, session_slot NULLS LAST, province, shop_name, id
       LIMIT 200`,
      carrier ? [like, carrier] : [like]
    )
    res.json({ ok: true, rows: r.rows, query: q })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── GET rows by month (YYYY-MM) ──
app.get('/api/rows/month', async (req, res) => {
  try {
    const { month } = req.query // format: 2026-03
    if (!month) return res.json({ ok: false, error: 'missing month' })
    const r = await pool.query(
      `SELECT * FROM delivery_rows
       WHERE TO_CHAR(session_date, 'YYYY-MM') = $1
       ORDER BY session_date, carrier, session_slot NULLS LAST, province, shop_name, id`,
      [month]
    )
    res.json({ ok: true, rows: r.rows, month })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── GET session dates (ย้อนหลัง 30 วัน) ──
app.get('/api/sessions', async (req, res) => {
  try {
    const r = await pool.query(
      `SELECT session_date, carrier, COUNT(*) as count
       FROM delivery_rows
       GROUP BY session_date, carrier
       ORDER BY session_date DESC
       LIMIT 60`
    )
    res.json({ ok: true, sessions: r.rows })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── INSERT row ──
app.post('/api/rows', async (req, res) => {
  try {
    const { carrier, shop_name, province, invoice_no, quantity, red_sticker, session_date, session_slot } = req.body
    const date = session_date || new Date().toISOString().slice(0, 10)
    const r = await pool.query(
      `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker, session_slot)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8) RETURNING *`,
      [date, carrier, shop_name||'', province||'', invoice_no||'', quantity||0, red_sticker||false, session_slot||null]
    )
    res.json({ ok: true, row: r.rows[0] })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── UPDATE row ──
app.put('/api/rows/:id', async (req, res) => {
  try {
    const { shop_name, province, invoice_no, quantity, red_sticker } = req.body
    const r = await pool.query(
      `UPDATE delivery_rows SET shop_name=$1, province=$2, invoice_no=$3, quantity=$4, red_sticker=$5, updated_at=NOW()
       WHERE id=$6 RETURNING *`,
      [shop_name||'', province||'', invoice_no||'', quantity||0, red_sticker||false, req.params.id]
    )
    res.json({ ok: true, row: r.rows[0] })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── DELETE row ──
app.delete('/api/rows/:id', async (req, res) => {
  try {
    await pool.query('DELETE FROM delivery_rows WHERE id=$1', [req.params.id])
    res.json({ ok: true })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── DELETE all rows (carrier + date) ──
app.delete('/api/rows', async (req, res) => {
  try {
    const { carrier, date } = req.query
    if (!carrier || !date) return res.json({ ok: false, error: 'missing params' })
    await pool.query('DELETE FROM delivery_rows WHERE carrier=$1 AND session_date=$2', [carrier, date])
    res.json({ ok: true })
  } catch (e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── OCR ──
app.post('/api/ocr', upload.array('images', 20), async (req, res) => {
  try {
    const files = req.files
    if (!files?.length) return res.json({ ok: false, error: 'ไม่มีรูป' })

    const carrier = req.body.carrier || 'inter'
    const session_date = req.body.session_date || new Date().toISOString().slice(0, 10)
    const session_slot = req.body.session_slot || null
    const results = []

    for (const file of files) {
      const base64 = file.buffer.toString('base64')
      try {
        let data = { shop_name: '', province: '', invoice_no: '', quantity: 0 }
        try {
          const parsed = await ocrViaGateway(base64, file.mimetype)
          data = { ...data, ...parsed }
        } catch (ocrErr) {
          console.error('OCR error:', ocrErr.message)
        }

        // บันทึกลง DB เลย
        const dbRow = await pool.query(
          `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker, session_slot)
           VALUES ($1,$2,$3,$4,$5,$6,false,$7) RETURNING *`,
          [session_date, carrier, data.shop_name, data.province, data.invoice_no, data.quantity, session_slot]
        )
        results.push(dbRow.rows[0])
      } catch (e) {
        console.error('OCR error:', e.message)
        const dbRow = await pool.query(
          `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker, session_slot)
           VALUES ($1,$2,'','','',0,false,$3) RETURNING *`,
          [session_date, carrier, session_slot]
        )
        results.push({ ...dbRow.rows[0], _error: 'อ่านไม่ได้' })
      }
    }

    res.json({ ok: true, results })
  } catch (err) {
    console.error(err)
    res.json({ ok: false, error: err.message })
  }
})

// ── Export Excel ──
app.post('/api/export', (req, res) => {
  const { rows, carrier, inter, dmk, date } = req.body
  const exportDate = date || new Date().toISOString().slice(0, 10)
  const colWidths = [{ wch: 7 }, { wch: 40 }, { wch: 16 }, { wch: 10 }, { wch: 20 }, { wch: 10 }]
  const makeSheet = (r) => {
    const data = [
      ['ลำดับ', 'ชื่อร้านค้า / โรงพยาบาล / คลินิก', 'จังหวัด', 'กล่อง', 'Invoice', 'คาดแดง'],
      ...r.map((row, i) => [i+1, row.shop_name||'', row.province||'', row.quantity||0, row.invoice_no||'', row.red_sticker ? 'คาดแดง' : ''])
    ]
    const ws = XLSX.utils.aoa_to_sheet(data)
    ws['!cols'] = colWidths
    return ws
  }

  const wb = XLSX.utils.book_new()

  // Export ทั้งหมด (Inter + DMK รวมกัน 2 sheet)
  if (inter || dmk) {
    if (inter?.length) XLSX.utils.book_append_sheet(wb, makeSheet(inter), '🚛 Inter')
    if (dmk?.length)   XLSX.utils.book_append_sheet(wb, makeSheet(dmk),   '🚚 DMK')
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''newlife_all_${exportDate}.xlsx`)
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return res.send(buf)
  }

  // Export แยก carrier
  if (!rows?.length) return res.status(400).send('ไม่มีข้อมูล')
  XLSX.utils.book_append_sheet(wb, makeSheet(rows), carrier === 'dmk' ? '🚚 DMK' : '🚛 Inter')
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''newlife_${carrier}_${exportDate}.xlsx`)
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  res.send(buf)
})

// ── Export Excel by Month ──
app.get('/api/export/month', async (req, res) => {
  try {
    const { month, carrier: filterCarrier } = req.query
    if (!month) return res.status(400).send('missing month')

    const r = await pool.query(
      `SELECT * FROM delivery_rows
       WHERE TO_CHAR(session_date, 'YYYY-MM') = $1
       ${filterCarrier ? 'AND carrier = $2' : ''}
       ORDER BY session_date, carrier, session_slot NULLS LAST, province, shop_name, id`,
      filterCarrier ? [month, filterCarrier] : [month]
    )

    const allRows = r.rows
    const interRows = allRows.filter(x => x.carrier === 'inter')
    const dmkRows   = allRows.filter(x => x.carrier === 'dmk')

    const colW = [{ wch: 7 }, { wch: 12 }, { wch: 40 }, { wch: 16 }, { wch: 10 }, { wch: 20 }, { wch: 10 }]
    const makeSheet = (rows) => {
      const data = [
        ['ลำดับ', 'วันที่', 'ชื่อร้านค้า / โรงพยาบาล / คลินิก', 'จังหวัด', 'กล่อง', 'Invoice', 'คาดแดง'],
        ...rows.map((row, i) => [
          i+1,
          String(row.session_date).slice(0,10),
          row.shop_name||'', row.province||'', row.quantity||0,
          row.invoice_no||'', row.red_sticker ? 'คาดแดง' : ''
        ])
      ]
      const ws = XLSX.utils.aoa_to_sheet(data)
      ws['!cols'] = colW
      return ws
    }

    const wb = XLSX.utils.book_new()
    if (filterCarrier === 'inter') {
      XLSX.utils.book_append_sheet(wb, makeSheet(interRows), '🚛 Inter')
    } else if (filterCarrier === 'dmk') {
      XLSX.utils.book_append_sheet(wb, makeSheet(dmkRows), '🚚 DMK')
    } else {
      if (interRows.length) XLSX.utils.book_append_sheet(wb, makeSheet(interRows), '🚛 Inter')
      if (dmkRows.length)   XLSX.utils.book_append_sheet(wb, makeSheet(dmkRows),   '🚚 DMK')
    }

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
    const label = filterCarrier || 'all'
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''newlife_${label}_${month}.xlsx`)
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.send(buf)
  } catch(e) {
    res.status(500).send(e.message)
  }
})

// ── Import from server-side file (one-time bulk import) ──
app.post('/api/import-file', async (req, res) => {
  try {
    const { filename, carrier, secret, delete_month } = req.body
    if (secret !== (process.env.IMPORT_SECRET || 'newlife2026')) return res.json({ ok: false, error: 'unauthorized' })

    // ลบ data เก่าถ้าต้องการ
    if (delete_month) {
      const del = await pool.query(
        `DELETE FROM delivery_rows WHERE TO_CHAR(session_date,'YYYY-MM')=$1 AND carrier=$2`,
        [delete_month, carrier || 'inter']
      )
      console.log(`Deleted ${del.rowCount} rows for ${delete_month}`)
    }

    const filePath = path.join(__dirname, filename)
    const fs = require('fs')
    if (!fs.existsSync(filePath)) return res.json({ ok: false, error: 'file not found: ' + filename })

    function parseSheet(name) {
      // ดึง slot ก่อน
      let slot = null
      if (name.includes('เช้า')) slot = 'เช้า'
      else if (name.includes('บ่าย')) slot = 'บ่าย'

      // รองรับหลาย pattern ของชื่อ sheet
      const thaiMonths = {
        'ม.ค.':  '01', 'ก.พ.': '02', 'มี.ค.': '03', 'เม.ย.': '04',
        'พ.ค.':  '05', 'มิ.ย.': '06', 'ก.ค.':  '07', 'ส.ค.':  '08',
        'ก.ย.':  '09', 'ต.ค.':  '10', 'พ.ย.':  '11', 'ธ.ค.':  '12'
      }
      for (const [abbr, mm] of Object.entries(thaiMonths)) {
        const m = name.match(new RegExp('(\\d+)\\s*' + abbr.replace('.','\\.')))
        if (m) {
          const day = String(m[1]).padStart(2, '0')
          return { date: `2026-${mm}-${day}`, slot }
        }
      }
      // fallback: ตัวเลข เช่น "10/3" หรือ "10-3"
      const numM = name.match(/(\d{1,2})[\/\-](\d{1,2})/)
      if (numM) {
        const day = String(numM[1]).padStart(2,'0')
        const mo  = String(numM[2]).padStart(2,'0')
        return { date: `2026-${mo}-${day}`, slot }
      }
      return null
    }

    const XLSX2 = require('xlsx')
    const wb = XLSX2.readFile(filePath)
    let total = 0, skipped = 0, log = []

    for (const sheetName of wb.SheetNames) {
      const parsed = parseSheet(sheetName)
      if (!parsed) { skipped++; continue }
      const { date, slot } = parsed

      const ws = wb.Sheets[sheetName]
      const rows = XLSX2.utils.sheet_to_json(ws, { header: 1, defval: '' })
      // กรองเฉพาะแถวที่มีชื่อร้านจริงๆ — ไม่เอา template row ว่าง
      const dataRows = rows.filter(r => typeof r[0] === 'number' && r[0] > 0 && String(r[1]||'').trim())
      if (!dataRows.length) { skipped++; continue }
      log.push(`${sheetName} → ${date}${slot ? ' ['+slot+']' : ''} : ${dataRows.length} rows`)

      for (const row of dataRows) {
        // แยก invoice กับ คาดแดง ออกจากกัน
        let rawNote = String(row[4]||'').trim()
        let red_sticker = false
        if (rawNote.includes('คาดแดง')) {
          red_sticker = true
          rawNote = rawNote.replace(/\s*คาดแดง\s*/g, '').trim()
        }
        await pool.query(
          `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker, session_slot)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8)`,
          [date, carrier || 'inter', String(row[1]||'').trim(), String(row[2]||'').trim(), rawNote, parseInt(row[3])||0, red_sticker, slot]
        )
        total++
      }
    }
    res.json({ ok: true, total, skipped, log })
  } catch(e) {
    res.json({ ok: false, error: e.message })
  }
})

// ── Import Excel ──
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    const { carrier, session_date } = req.body
    if (!req.file) return res.json({ ok: false, error: 'ไม่มีไฟล์' })
    if (!carrier || !session_date) return res.json({ ok: false, error: 'missing carrier/date' })

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' })
    const ws = wb.Sheets[wb.SheetNames[0]]
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' })

    // หา header row (row ที่มี ชื่อ / ลำดับ)
    let dataRows = raw.filter((row, i) => i > 0 && row.some(c => String(c).trim()))
    // ลอง detect columns จาก header
    const header = raw[0]?.map(c => String(c).toLowerCase().trim()) || []
    const colIdx = {
      shop:    header.findIndex(h => h.includes('ชื่อ') || h.includes('shop')),
      province:header.findIndex(h => h.includes('จังหวัด') || h.includes('province')),
      qty:     header.findIndex(h => h.includes('กล่อง') || h.includes('qty') || h.includes('quantity') || h.includes('ลัง')),
      invoice: header.findIndex(h => h.includes('invoice') || h.includes('หมายเหตุ')),
      red:     header.findIndex(h => h.includes('คาดแดง') || h.includes('red')),
    }
    // fallback index ถ้าไม่เจอ header
    if (colIdx.shop < 0)     colIdx.shop = 1
    if (colIdx.province < 0) colIdx.province = 2
    if (colIdx.qty < 0)      colIdx.qty = 3
    if (colIdx.invoice < 0)  colIdx.invoice = 4
    if (colIdx.red < 0)      colIdx.red = 5

    const inserted = []
    for (const row of dataRows) {
      const shop_name  = String(row[colIdx.shop]   || '').trim()
      const province   = String(row[colIdx.province]|| '').trim()
      const invoice_no = String(row[colIdx.invoice] || '').trim()
      const quantity   = parseInt(row[colIdx.qty])  || 0
      const red_sticker = String(row[colIdx.red]||'').includes('คาด') || String(row[colIdx.red]||'').toLowerCase().includes('red')

      if (!shop_name && !invoice_no) continue // skip empty rows

      const r = await pool.query(
        `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker)
         VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING *`,
        [session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker]
      )
      inserted.push(r.rows[0])
    }

    res.json({ ok: true, inserted, count: inserted.length })
  } catch(e) {
    console.error('import error:', e)
    res.json({ ok: false, error: e.message })
  }
})

// ── Update API Key at runtime (no redeploy needed) ──
app.post('/api/update-key', (req, res) => {
  const { key, secret } = req.body
  if (secret !== (process.env.IMPORT_SECRET || 'newlife2026')) return res.json({ ok: false, error: 'unauthorized' })
  if (!key || !key.startsWith('sk-ant-')) return res.json({ ok: false, error: 'invalid key format' })
  process.env.ANTHROPIC_API_KEY = key
  console.log(`🔑 API key updated at ${new Date().toISOString()}`)
  res.json({ ok: true })
})

// ── START ──
async function startServer() {
  app.listen(PORT, () => console.log(`✅ NewLife Delivery OCR → http://localhost:${PORT}`))
  // retry DB init ให้ postgres มีเวลา boot
  for (let i = 1; i <= 10; i++) {
    try {
      await initDB()
      break
    } catch(e) {
      console.log(`⏳ DB not ready (attempt ${i}/10): ${e.message}`)
      await new Promise(r => setTimeout(r, 3000))
    }
  }
}
startServer()
