const XLSX = require('xlsx')
const _ = require('lodash')

const createDupandUniq = (workbook) => {
    return new Promise((resolve, reject) => {
        var wbExp

        const first_sheet_name = workbook.SheetNames[0]

        const worksheet = workbook.Sheets[first_sheet_name]

        const jsonObj = XLSX.utils.sheet_to_json(worksheet)

        const objKeys = Object.keys(jsonObj[0])

        if (objKeys.length > 4) {
            return reject('File must contain 4 or less competitors')
        }

        var CombArr = []
        var obje = {}
        var competitors = objKeys
        var isHeader

        objKeys.forEach(x => {
            if (x == '__EMPTY')
                isHeader = true
        })

        if (isHeader) {
            competitors = []
            objKeys.forEach(x => {
                competitors = [...competitors, jsonObj[0][x]]
            })
            delete jsonObj[0]
        }

        //Create an array with all duplicated keywords removed from the same column
        objKeys.forEach(obj => {
            var arr = []
            jsonObj.forEach(x => {
                if (x[obj]) {
                    arr.push(x[obj])
                }
            })
            arr = arr.map(x => {
                if (typeof x === 'string') {
                    return x.toLowerCase()
                } return x
            })
            arr = _.uniq(arr)
            obje[obj] = [...arr]
            CombArr = [...CombArr, ...arr]
        })
        console.log('obje', obje)

        //Find duplicate and unique keys
        const findDuplicateAndUniq = (arr, callback) => {
            var object = {}
            var duplicate = []
            var duplicatee = []
            var duplicateCountwise = []
            var dupInHowMany = []
            var a = []
            var unique = []

            arr.forEach(x => {
                if (!object[x]) {
                    object[x] = 0
                }
                object[x] += 1
            })

            uniquer = Object.keys(object).filter(x => object[x] < 2)
            console.log('uniquer', uniquer)

            var competitorsLength = competitors.length
            console.log(competitorsLength)
            for (i = 2; i < competitorsLength; i++) {
                let str = 'Common keyword(s) in any ' + i
                dupInHowMany = [...dupInHowMany, str]
            }

            dupInHowMany = [...dupInHowMany, 'Common keyword(s) in all']

            for (var i = 2; i <= competitorsLength; i++) {
                duplicateCountwise = [...duplicateCountwise, Object.keys(object).filter(x => object[x] == i)]
            }

            duplicateCountwise.map(x => {
                if (x.length == 0) {
                    return x[0] = 'NO_ANY_DUPLICATES'
                }
            })

            dupVer = []
            var longLength = 0
            duplicateCountwise.forEach(x => {
                if (x.length > longLength) {
                    return longLength = x.length
                }
            })

            for (var i = 0; i < longLength; i++) {
                dupVer[i] = []
                duplicateCountwise.forEach(x => {
                    if (!x[i]) {
                        return dupVer[i] = [...dupVer[i], undefined]
                    }
                    dupVer[i] = [...dupVer[i], x[i]]
                })
            }

            dupVer = [dupInHowMany, ...dupVer]

            Object.keys(obje).forEach(x => obje[x] = obje[x].filter(f => uniquer.includes(f.toString())))

            Object.keys(obje).forEach(x => {
                if (obje[x].length == 0) {
                    return obje[x][0] = 'NO_UNIQUE_KEYWORDS'
                }
            })

            var verarr = []
            var long = 0
            Object.keys(obje).forEach(o => {
                if (long < obje[o].length) { long = obje[o].length }
            })

            for (var i = 0; i < long; i++) {
                var temp = []
                Object.keys(obje).forEach(o => {
                    if (!obje[o][i]) {
                        return temp = [...temp, undefined]
                    }
                    temp = [...temp, obje[o][i]]
                })
                verarr = [...verarr, temp]
            }

            callback(dupVer, [competitors, ...verarr])

        }

        // Get duplicate and unique keywords and write them to file

        findDuplicateAndUniq(CombArr, (dupData, uniqData) => {
            var ws_name = "Duplicate Keys"
            var ws1_name = "Unique Keys"

            if (typeof console !== 'undefined') console.log(new Date())
            var wb = XLSX.utils.book_new(), ws = XLSX.utils.aoa_to_sheet(dupData); ws1 = XLSX.utils.aoa_to_sheet(uniqData)

            /* add worksheet to workbook */
            XLSX.utils.book_append_sheet(wb, ws, ws_name)
            XLSX.utils.book_append_sheet(wb, ws1, ws1_name)


            if (!wb.Props) wb.Props = {}
            wb.Props.Title = "Duplicate And Unique Keywords"

            noOfColsWs = dupData[0].length

            ws['!cols'] = []

            for (var i = 0; i < noOfColsWs; i++) {
                ws['!cols'] = [...ws['!cols'], { wpx: 200 }]
            }

            noOfColsWs1 = uniqData[0].length

            ws1['!cols'] = []

            for (var i = 0; i < noOfColsWs1; i++) {
                ws1['!cols'] = [...ws['!cols'], { wpx: 200 }]
            }
            /* write workbook */
            if (typeof console !== 'undefined') console.log(new Date())
          
            wbExp = wb
            

            if (typeof console !== 'undefined') console.log(new Date())
        })

        return resolve(wbExp)

    })
}

module.exports = createDupandUniq