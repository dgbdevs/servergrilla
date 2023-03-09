 const XLSX = require('xlsx')
 const express = require('express')
const morgan = require('morgan')
const app = express()
const cors = require('cors');


app.use(cors({
    origin: '*'
}));

//settings 
app.set('port', process.env.PORT || 3000)
app.set('json spaces', 2)

//middlewares
app.use(morgan('dev'));
app.use(express.urlencoded({extended:false}));
app.use(express.json());

//routes

app.get('/',(req,res)=>{
    res.send('hello World')
})


app.get('/Jan',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[0]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})
app.get('/Feb',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[1]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})
app.get('/Mar',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[2]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})
app.get('/Apr',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[3]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})
app.get('/May',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[4]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})
app.get('/Jun',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[5]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})

app.get('/Jul',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[6]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})

app.get('/Aug',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[7]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})

app.get('/Sep',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[8]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})

app.get('/Oct',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[9]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})

app.get('/Nov',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[10]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})

app.get('/Dec',(req,res)=>{
    
    function leerExcel(ruta){

        const workbook = XLSX.readFile(ruta)
        const workbookSheets = workbook.SheetNames;
        const sheet = workbookSheets[11]
        const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])
    
        return dataExcel
    
    }    
     
    res.json(leerExcel('./Presencial_NOC-TR.xlsx'))
})

app.listen(app.get('port'),()=>{
    console.log(`Server on Port ${app.get('port')}`);
})

/* 
function leerExcel(ruta){

    const workbook = XLSX.readFile(ruta)
    const workbookSheets = workbook.SheetNames;
    const sheet = workbookSheets[0]
    const dataExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheet])

    console.log(dataExcel)

}

leerExcel('./Presencial_NOC-TR.xlsx') */