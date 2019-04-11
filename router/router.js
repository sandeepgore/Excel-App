const express = require('express')
const multer = require('multer')
const path = require('path')
const XLSX = require('xlsx')

const createDupandUniq = require('../src/xlsx')

const router = new express.Router()

const upload = multer({
    fileFilter(req, file, cb) {
        if (!file.originalname.match(/\.(xls|xlsx)$/)) {
            cb(new Error('Please upload Excel file'))
        }

        cb('', true)
    }
}).single('excel')

router.post('/upload/excel', async (req, res) => {
    upload(req, res, async function (err) {
        if (err instanceof multer.MulterError) {
            return res.status(400).send()

        } else if (err) {
            return res.render('error', { errorCode: 400, error: 'Please upload an xls | xlsx file' })

        }
        try {
            const buffer = await req.file.buffer

            const workbook = XLSX.read(buffer, { type: "buffer" })
            const wbExp = await createDupandUniq(workbook)
            if (wbExp) {
                // console.log(resObj)
                XLSX.writeFile(wbExp, "DupAndUniqKeys.xlsx")
                res.download("DupAndUniqKeys.xlsx")
            }
        } catch (error) {
            console.log(error)
            res.render('error', { errorCode: '400', error })
        }
    }, (error, req, res, next) => {
        res.render('error', {
            errorCode: 500,
            error: 'Internal Server Error!'
        })
    })
})

router.get('*', (req, res) => {
    res.render('error', {
        errorCode: 404,
        error: 'Page Not Found'
    })
})

module.exports = router