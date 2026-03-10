const express = require('express')
const multer = require('multer')
const path = require('path')
const Anthropic = require('@anthropic-ai/sdk')
const XLSX = require('xlsx')
const { Pool } = require('pg')

const app = express()
const PORT = process.env.PORT || 3100

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } })
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

// ── DB ──
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
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
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    )
  `)
  console.log('✅ DB ready')
}

app.use(express.json())
app.use(express.static(path.join(__dirname, 'public')))

// ── GET rows (วันนี้) ──
app.get('/api/rows', async (req, res) => {
  try {
    const date = req.query.date || new Date().toISOString().slice(0, 10)
    const r = await pool.query(
      `SELECT * FROM delivery_rows WHERE session_date = $1 ORDER BY carrier, sort_order, id`,
      [date]
    )
    res.json({ ok: true, rows: r.rows, date })
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
    const { carrier, shop_name, province, invoice_no, quantity, red_sticker, session_date } = req.body
    const date = session_date || new Date().toISOString().slice(0, 10)
    const r = await pool.query(
      `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker)
       VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING *`,
      [date, carrier, shop_name||'', province||'', invoice_no||'', quantity||0, red_sticker||false]
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
    const results = []

    for (const file of files) {
      const base64 = file.buffer.toString('base64')
      try {
        const msg = await anthropic.messages.create({
          model: 'claude-opus-4-5',
          max_tokens: 1024,
          messages: [{
            role: 'user',
            content: [
              { type: 'image', source: { type: 'base64', media_type: file.mimetype, data: base64 } },
              {
                type: 'text',
                text: `อ่านข้อมูลจากใบนำส่งสินค้าของบริษัทนิวไลฟ์ ฟาร์มา ในรูปนี้ แล้วตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่น:
{
  "shop_name": "ชื่อผู้รับสินค้าปลายทาง เช่น ห้างหุ้นส่วนสามัญ ร้านอุดมเวชภัณฑ์ (คัดลอกมาเต็มๆ)",
  "province": "ชื่อจังหวัดเท่านั้น ไม่มีคำว่า จังหวัด นำหน้า เช่น พิจิตร, กรุงเทพมหานคร",
  "invoice_no": "เลข Invoice หรือ BILL NO. เช่น IV69030970",
  "quantity": จำนวนรวม กล่องหรือลัง จากช่อง รวม ในใบนำส่ง (ตัวเลขจำนวนเต็ม ถ้าไม่พบใส่ 0)
}
ถ้าไม่พบข้อมูลให้ใส่ "" หรือ 0 ตามประเภท`
              }
            ]
          }]
        })

        let data = { shop_name: '', province: '', invoice_no: '', quantity: 0 }
        const text = msg.content[0].text.trim()
        const m = text.match(/\{[\s\S]*\}/)
        if (m) { try { data = { ...data, ...JSON.parse(m[0]) } } catch {} }

        // บันทึกลง DB เลย
        const dbRow = await pool.query(
          `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker)
           VALUES ($1,$2,$3,$4,$5,$6,false) RETURNING *`,
          [session_date, carrier, data.shop_name, data.province, data.invoice_no, data.quantity]
        )
        results.push(dbRow.rows[0])
      } catch (e) {
        console.error('OCR error:', e.message)
        // บันทึก row ว่างเปล่าถ้า OCR fail
        const dbRow = await pool.query(
          `INSERT INTO delivery_rows (session_date, carrier, shop_name, province, invoice_no, quantity, red_sticker)
           VALUES ($1,$2,'','','',0,false) RETURNING *`,
          [session_date, carrier]
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
  const { rows, carrier } = req.body
  if (!rows?.length) return res.status(400).send('ไม่มีข้อมูล')

  const carrierLabel = carrier === 'dmk' ? 'DMK' : 'Inter'
  const data = [
    ['ลำดับ', 'ชื่อ', 'จังหวัด', 'กล่อง', 'หมายเหตุ (Invoice)', 'คาดแดง'],
    ...rows.map((r, i) => [i+1, r.shop_name||'', r.province||'', r.quantity||0, r.invoice_no||'', r.red_sticker ? 'คาดแดง' : ''])
  ]

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet(data)
  ws['!cols'] = [{ wch: 7 }, { wch: 38 }, { wch: 16 }, { wch: 10 }, { wch: 18 }, { wch: 10 }]
  XLSX.utils.book_append_sheet(wb, ws, `ใบนำส่ง ${carrierLabel}`)

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
  const date = new Date().toISOString().slice(0, 10)
  res.setHeader('Content-Disposition', `attachment; filename="newlife_${carrier}_${date}.xlsx"`)
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  res.send(buf)
})

// ── START ──
initDB().then(() => {
  app.listen(PORT, () => console.log(`✅ NewLife Delivery OCR → http://localhost:${PORT}`))
}).catch(e => {
  console.error('DB init failed:', e.message)
  // start anyway แม้ไม่มี DB (fallback)
  app.listen(PORT, () => console.log(`⚠️ NewLife Delivery (no DB) → http://localhost:${PORT}`))
})
