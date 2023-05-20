function saveImageToDrive(postId, imageUrl, folderId) {
  const folder = DriveApp.getFolderById(folderId)
  const fileName = postId

  const response = UrlFetchApp.fetch(imageUrl)
  const oImageBlob = response.getBlob()
  const tmpFile = folder.createFile(oImageBlob.setName('tmp'))
  const link = Drive.Files.get(tmpFile.getId()).thumbnailLink.replace(/\=s.+/, "=s" + 500)

  const imageBlob = UrlFetchApp.fetch(link).getBlob()
  const file = folder.createFile(imageBlob.setName(fileName))

  Drive.Files.remove(tmpFile.getId())

  return file.getId()
}

function convertImages(uSheet) {
  const imageRange = uSheet.getRange("B:B")
  const imagesUrl = imageRange.getFormulas().map(row => {
    const matched = row[0].match(/"([^"]+)"/g)
    if (matched && matched.length) {
      return [matched[0].replace(/"/g, "")]
    } else {
      return [""]
    }
  })
  imageRange.setValues(imagesUrl)
}

function uploadToLibrary() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const uSheet = spreadsheet.getSheetByName('upload')
  const iSheet = spreadsheet.getSheetByName('images')
  convertImages(uSheet)
  const uRange = uSheet.getRange("A:C")

  const uValues = uRange.getDisplayValues()
  const iValues = iSheet.getRange("A:B").getValues()
  const folderId = uSheet.getRange("D1").getValue()

  // 清空结果列
  uSheet.getRange("C:C").clearContent()

  const uploaded = []
  for (let i = 0; i < uValues.length; i++) {
    const oId = uValues[i][0].toString().trim()

    if (oId && oId.startsWith('#') && uValues[i][1].toString().trim()) {
      const res = iValues.filter(iRow => iRow[0] === oId)
      if (!res.length) {
        const fileId = saveImageToDrive(oId.slice(1), uValues[i][1], folderId)
        if (fileId) {
          uploaded.push([oId, fileId])
          uValues[i][2] = '上传成功'
        }
      } else {
        uValues[i][2] = '已经存在'
      }
    }
  }

  if (uploaded.length) {
    iSheet.insertRows(1, uploaded.length)
    iSheet.getRange(1, 1, uploaded.length, 2).setValues(uploaded)
  }
  uRange.setValues(uValues)
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("📤🗃️").addItem('上传图片', 'uploadToLibrary').addToUi()
}