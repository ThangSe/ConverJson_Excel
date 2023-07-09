const express = require("express")
const excelJs = require("exceljs")
const bodyParser = require('body-parser')
require('dotenv').config()
var fs = require("fs")

const app = express()

const PORT = process.env.PORT || 3000
app.use(bodyParser.json({limit:"50mb"}))
app.use(bodyParser.urlencoded({ extended: true }))
app.post("/JsonExportExcel/Convert-All", async (req, res) => {
    try {
        let object
        let workbook = new excelJs.Workbook()
        let row = 1
        try{
            workbook = await workbook.xlsx.readFile(req.body.ExcelPath);
        }   
        catch (err){
            throw new Error("Excel file not found at path " + req.body.ExcelPath)        
        }
        
        const sheet = workbook.addWorksheet(req.body.SheetName)
        sheet.columns = [
            {header: "Id", key: "id", width: 5},
            {header: "FileName", key: "filename", width: 25},
            {header: "Image", key: 'image', width: 15},
            {header: "Object", key: "object", width: 25},
            {header: "Category", key: "category", width: 25},
            {header: "Type", key: "type", width: 25},
        ]  
        try{
            object = JSON.parse(fs.readFileSync(req.body.ConfigJsonPath, 'utf8'))
        }
        catch(err) {
            throw new Error("Config file not found at path " + req.body.ConfigJsonPath)
        }
        await object.map((value, index) => {
            const imageId = workbook.addImage({
                filename: `${req.body.ImagePath}${value.FileName}`,
                extension: 'png',
            })
            sheet.addRow({
                id: value.Id,
                filename: value.FileName,
                object: value.object,
                type: value.type,
                category: value.Category,
            })
            sheet.addImage(imageId, {
                tl: {col: sheet.getColumn("image").number - 1, row: row},
                ext: {width: 100, height: 20},
                editAs: 'undefined'
            })
            row++
        })
        await workbook.xlsx.writeFile(req.body.ExcelPath);
        res.status(200).json("Success");

    } catch (err) {
        res.status(500).json(err.message)
    }
})

app.post("/JsonExportExcel/Edit-Exist", async (req, res) => {
    try {
        let object
        let sheet
        var nextPosDataRow = req.body.StartPosDataRow
        let workbook = new excelJs.Workbook()
        try{
            workbook = await workbook.xlsx.readFile(req.body.ExcelPath);
        }   
        catch (err){
            throw new Error("Excel file not found at path " + req.body.ExcelPath)        
        }
        try {
            sheet = await workbook.getWorksheet(req.body.SheetName)
            if(!sheet) throw new Error()
        } 
        catch (err) {
            throw new Error("Sheet name " + req.body.SheetName + " does not existed")
        }

        const rowCount = sheet.rowCount
        sheet.columns = [
            {header: "Id", key: "id", width: 5},
            {header: "FileName", key: "filename", width: 25},
            {header: "Image", key: 'image', width: 15},
            {header: "Object", key: "object", width: 25},
            {header: "Category", key: "category", width: 25},
            {header: "Type", key: "type", width: 25},
        ]  
        try{
            object = JSON.parse(fs.readFileSync(req.body.ConfigJsonPath, 'utf8'))
        }
        catch(err) {
            throw new Error("Config file not found at path " + req.body.ConfigJsonPath)
        }   
        await object.map((value) => {
            const imageId = workbook.addImage({
                filename: `${req.body.ImagePath}${value.FileName}`,
                extension: 'png',
            })
            for (var i = nextPosDataRow; i <= rowCount; i++) {
                const row = sheet.getRow(i)
                if(row.values[1] == value.Id) {
                    row.values = {
                        id: value.Id,
                        filename: value.FileName,
                        object: value.object,
                        type: value.type,
                        category: value.Category,       
                    }
                    sheet.addImage(imageId, {
                        tl: {col: sheet.getColumn("image").number - 1, row: row.number - 1},
                        ext: {width: 100, height: 20},
                        editAs: 'undefined'
                    })
                    nextPosDataRow++
                    break
                }
            }
        })
        await workbook.xlsx.writeFile(req.body.ExcelPath);
        res.status(200).json("Success");

    } catch (err) {
        res.status(500).json(err.message)
    }
})

app.listen(PORT, () => {
    console.log(`Server Running on PORT: ${PORT}`)
})