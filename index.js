const express = require("express")
const excelJs = require("exceljs")
const bodyParser = require('body-parser')
require('dotenv').config()
var fs = require("fs")

const app = express()

const PORT = process.env.PORT || 4000
app.use(bodyParser.json({limit:"50mb"}))
app.use(bodyParser.urlencoded({ extended: true }))
app.get("/JsonExportExcel", async (req, res) => {
    try {
        let object
        let workbook = new excelJs.Workbook()
        let row = req.body.ImageStartPos.split('')[0]
        let column = req.body.ImageStartPos.split('')[1]
        try{
            workbook = await workbook.xlsx.readFile(req.body.ExcelPath);
        }   
        catch (err){
            throw new Error("Excel file not found " + req.body.ExcelPath)        
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
            throw new Error("Config file not found " + req.body.ConfigJsonPath)
        }
        await object.map((value) => {
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
            sheet.addImage(imageId, `${row}${column}:${row}${column}`)
            column++;
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