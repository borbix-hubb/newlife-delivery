const express = require('express')
const multer = require('multer')
const path = require('path')
const Anthropic = require('@anthropic-ai/sdk')
const XLSX = require('xlsx')

const app = express()
const PORT = 3100

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } })
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })

app.use(express.json())
app.use(express.static(path.join(__dirname, 'public')))

// OCR — รองรับหลายรูปพร้อมกัน
app.post('/api/ocr', upload.array('images', 20), async (req, res) => {
  try {
    const files = req.files
    if (!files || !files.length) return res.json({ ok: false, error: 'ไม่มีรูป' })

    const results = []

    for (let idx = 0; idx < files.length; idx++) {
      const file = files[idx]
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
        if (m) {
          try { data = { ...data, ...JSON.parse(m[0]) } } catch {}
        }
        results.push({ ...data, red_sticker: false, _file: file.originalname })
      } catch (e) {
        results.push({ shop_name: '', province: '', invoice_no: '', quantity: 0, note: 'อ่านไม่ได้', _file: file.originalname })
      }
    }

    res.json({ ok: true, results })
  } catch (err) {
    console.error(err)
    res.json({ ok: false, error: err.message })
  }
})

// Export Excel
app.post('/api/export', (req, res) => {
  const { rows, carrier } = req.body
  if (!rows?.length) return res.status(400).send('ไม่มีข้อมูล')

  const carrierLabel = carrier === 'dmk' ? 'DMK' : 'Inter'
  const sheetName = `ใบนำส่ง ${carrierLabel}`

  const data = [
    ['ลำดับ', 'ชื่อ', 'จังหวัด', 'กล่อง', 'หมายเหตุ (Invoice)', 'คาดแดง'],
    ...rows.map((r, i) => [
      i + 1,
      r.shop_name || '',
      r.province || '',
      r.quantity || 0,
      r.invoice_no || '',
      r.red_sticker ? 'คาดแดง' : ''
    ])
  ]

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.aoa_to_sheet(data)
  ws['!cols'] = [{ wch: 7 }, { wch: 38 }, { wch: 16 }, { wch: 10 }, { wch: 18 }, { wch: 10 }]
  XLSX.utils.book_append_sheet(wb, ws, sheetName)

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
  const date = new Date().toISOString().slice(0, 10)
  res.setHeader('Content-Disposition', `attachment; filename="newlife_${carrier}_${date}.xlsx"`)
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  res.send(buf)
})

app.listen(PORT, () => console.log(`✅ NewLife Delivery OCR → http://localhost:${PORT}`))
