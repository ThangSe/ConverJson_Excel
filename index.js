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
        let objects
        let workbook = new excelJs.Workbook()
        let row = 1
        try{
            workbook = await workbook.xlsx.readFile(req.body.ExcelPath);
        }   
        catch (err){
            throw new Error("Excel file not found at path " + req.body.ExcelPath)        
        }
        
        const sheet = workbook.addWorksheet(req.body.SheetName)
        let columns = new Array()
        req.body.ColumnsToConVert.forEach(column => {
            let width = 25
            if(column === "Id") width = 5
            columns.push({header: `${column}`, key: `${column.toLowerCase()}`, width: width})
        })

        sheet.columns = columns 

        try{
            objects = JSON.parse(fs.readFileSync(req.body.ConfigJsonPath, 'utf8'))
        }
        catch(err) {
            throw new Error("Config file not found at path " + req.body.ConfigJsonPath)
        }
        
        await objects.map((object) => {
            const objectRow = {}
            const imageId = workbook.addImage({
                filename: `${req.body.ImagePath}${object.FileName}`,
                extension: 'png',
            })
            req.body.ColumnsToConVert.forEach(column => {
                objectRow[column.toLowerCase()] = object[column]
            })
            sheet.addRow(objectRow)
            sheet.addImage(imageId, {
                tl: {col: sheet.getColumn("image").number - 1, row: row},
                ext: {width: 100, height: 20},
                editAs: 'undefined',
                /*hyperlinks: {
                    hyperlink: 'https://www.google.com.vn/?hl=vi',
                }*/
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
        let objects
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
        let columns = new Array()
        req.body.ColumnsToConVert.forEach(column => {
            let width = 25
            if(column === "Id") width = 5
            columns.push({header: `${column}`, key: `${column.toLowerCase()}`, width: width})
        })

        sheet.columns = columns 
        try{
            objects = JSON.parse(fs.readFileSync(req.body.ConfigJsonPath, 'utf8'))
        }
        catch(err) {
            throw new Error("Config file not found at path " + req.body.ConfigJsonPath)
        }   
        await objects.map((object) => {
            const imageId = workbook.addImage({
                filename: `${req.body.ImagePath}${object.FileName}`,
                extension: 'png',
            })
            for (var i = nextPosDataRow; i <= rowCount; i++) {
                const row = sheet.getRow(i)
                const objectRow = {}
                if(row.values[1] == object.Id) {
                    req.body.ColumnsToConVert.forEach(column => {
                        objectRow[column.toLowerCase()] = object[column]
                    })
                    row.values = objectRow
                    sheet.addImage(imageId, {
                        tl: {col: sheet.getColumn("image").number - 1, row: row.number - 1},
                        ext: {width: 100, height: 20},
                        editAs: 'undefined',
                        /*hyperlinks: {
                            hyperlink: 'https://www.google.com.vn/?hl=vi',
                          }*/
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