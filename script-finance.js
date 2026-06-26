const HEADERS = {
  Users:         ['id','username','password','role','createdAt'],
  Assets:        ['id','userId','category','name','amount','note','updatedAt'],
  Debts:         ['id','userId','name','amount','rate','term','startDate','note','updatedAt'],
  Income:        ['id','userId','name','amount','category','month','year','note','updatedAt'],
  Expenses:      ['id','userId','name','amount','category','month','year','note','updatedAt'],
  Investments:   ['id','userId','name','type','buyPrice','currentPrice','buyDate','note','updatedAt'],
  Goals:         ['id','userId','name','targetAmount','currentAmount','deadline','note','updatedAt'],
  BusinessPlans: ['id','userId','name','startDate','status','revenue','cost','profit','note','updatedAt'],
  Transactions:  ['id','userId','type','name','amount','category','date','note']
}

function doGet(e) {
  try {
    const sheet = e.parameter.sheet || 'all'
    const userId = e.parameter.userId
    let result = {}
    const sheets = Object.keys(HEADERS).filter(s => s !== 'Users')
    if (sheet === 'all') {
      sheets.forEach(s => {
        let rows = readSheet(s)
        if (userId) rows = rows.filter(r => String(r.userId) === String(userId))
        result[s] = rows
      })
    } else {
      let rows = readSheet(sheet)
      if (userId) rows = rows.filter(r => String(r.userId) === String(userId))
      result[sheet] = rows
    }
    return ok(result)
  } catch(err) { return err_(err) }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents)
    const { action, sheet, data } = body
    let result

    if (action === 'login') {
      const users = readSheet('Users')
      const user = users.find(u => u.username === data.username && u.password === data.password)
      if (user) return ok({ user: { id: user.id, username: user.username, role: user.role } })
      else return ok_raw({ success: false, error: 'Sai tên đăng nhập hoặc mật khẩu' })
    }

    if (action === 'createUser') {
      const existing = readSheet('Users').find(u => u.username === data.username)
      if (existing) return ok_raw({ success: false, error: 'Username đã tồn tại' })
      result = upsertRow('Users', data)
      return ok(result)
    }

    if (action === 'upsert')       result = upsertRow(sheet, data)
    else if (action === 'delete')  result = deleteRow(sheet, data.id)
    else if (action === 'batchUpsert') result = data.map(row => upsertRow(sheet, row))

    return ok(result)
  } catch(err) { return err_(err) }
}

// ── Helpers ──────────────────────────────────────────────────────────────
function ok(data)     { return ContentService.createTextOutput(JSON.stringify({success:true,data})).setMimeType(ContentService.MimeType.JSON) }
function ok_raw(obj)  { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON) }
function err_(e)      { return ContentService.createTextOutput(JSON.stringify({success:false,error:e.message})).setMimeType(ContentService.MimeType.JSON) }

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName(sheetName)
  if (!sheet) {
    sheet = ss.insertSheet(sheetName)
    const headers = HEADERS[sheetName]
    if (headers) {
      sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold')
      sheet.setFrozenRows(1)
    }
  }
  return sheet
}

function readSheet(sheetName) {
  const sheet = getOrCreateSheet(sheetName)
  const data = sheet.getDataRange().getValues()
  if (data.length <= 1) return []
  const headers = data[0]
  return data.slice(1).map(row => {
    const obj = {}; headers.forEach((h,i) => { obj[h] = row[i] }); return obj
  }).filter(row => row.id && row.id !== '')
}

function upsertRow(sheetName, rowData) {
  const sheet = getOrCreateSheet(sheetName)
  const headers = HEADERS[sheetName]
  if (sheetName !== 'Users') rowData.updatedAt = new Date().toISOString()
  const data = sheet.getDataRange().getValues()
  let targetRow = -1
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rowData.id)) { targetRow = i+1; break }
  }
  const values = [headers.map(h => rowData[h] !== undefined ? rowData[h] : '')]
  if (targetRow > 0) sheet.getRange(targetRow,1,1,headers.length).setValues(values)
  else sheet.appendRow(values[0])
  return rowData
}

function deleteRow(sheetName, id) {
  const sheet = getOrCreateSheet(sheetName)
  const data = sheet.getDataRange().getValues()
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return {deleted:true,id} }
  }
  return {deleted:false,id}
}

function setupAllSheets() {
  Object.keys(HEADERS).forEach(name => getOrCreateSheet(name))
  // Tạo admin mặc định nếu chưa có
  const users = readSheet('Users')
  if (!users.find(u => u.username === 'admin')) {
    upsertRow('Users', {
      id: 'admin',
      username: 'admin',
      password: 'admin123',
      role: 'admin',
      createdAt: new Date().toISOString()
    })
  }
  SpreadsheetApp.getUi().alert('✅ Setup xong! Admin: admin / admin123')
}