const xlsx=require('exceljs')
const path = require("path")

module.exports.readWorkbook=function(filePath){
    let workbook=new xlsx.Workbook()
    let targetWorkbook=new xlsx.Workbook()
    if(filePath!=null){
        dirPath=path.dirname(filePath)
        let targetFile=path.join(dirPath,"tbresult"+(new Date()).toISOString()+".xlsx")
        workbook.xlsx.readFile(filePath).then(function(){
            let worksheet=workbook.getWorksheet("Sheet1")
            if(worksheet!=null||worksheet!=undefined){
                  //add two summary sheet
              let summarySheet=workbook.addWorksheet('Summary')
              let subSummarySheet=workbook.addWorksheet("subsummary")
              summarySheet.addRow(worksheet.getRow(2).values)
              subSummarySheet.addRow(worksheet.getRow(2).values)
      
              //get month
              let titleRow=worksheet.getRow(1)
              let monthCellValue=titleRow.getCell(3).value.toString()
              let monthPoint=monthCellValue.indexOf("--")
              let monthValue=monthCellValue.slice(monthPoint+2).trim()
              let monthArr=monthValue.split(".")
              if(monthArr[1]!=null){
                  if(monthArr[1].length==1){
                  monthArr[1]="0"+monthArr[1]
                  }
              }
              monthValue=monthArr[0]+monthArr[1]+'01'
              //remove first and second row
              worksheet.spliceRows(1,2)
              let rowLen=worksheet.rowCount      
              worksheet.spliceRows(rowLen,1)
      
              rowLen=worksheet.rowCount
              //format categoryCode and CategoryName     
              for(let rowIndex=1;rowIndex<=rowLen;rowIndex++){
                  let row=worksheet.getRow(rowIndex)
                  if(row.getCell(1).value!=null){
                  row.getCell(1).value=row.getCell(1).value.toString().trim()
                  }
                  if(row.getCell(2).value!=null){
                  row.getCell(2).value=row.getCell(2).value.toString().trim()
                  }
              }
              //Map Parent Name to subordinate
              for(let rowIndex=1;rowIndex<=rowLen;rowIndex++){
                  let row=worksheet.getRow(rowIndex)
                  let categoryCode=null
                  let categoryName=null
                  if(row.getCell(1).value!=null){
                  categoryCode=row.getCell(1).value.toString().trim()
                  row.getCell(1).value=categoryCode
                  } else{
                  categoryCode=null
                  }
                  if(row.getCell(2).value!=null){
                  categoryName=row.getCell(2).value.toString().trim()
                  row.getCell(2).value=categoryName
                  } else{
                  categoryName=null
                  }
                  if(categoryCode!=null){
                  let slicePoint=categoryCode.indexOf('.')
                  let sliceLastPoint=categoryCode.lastIndexOf('.')
                  if(categoryCode.length==4 && slicePoint<0){
                      MaptoSubCategory(categoryCode,categoryName,rowIndex,rowLen,worksheet)
                  }
                  if(slicePoint>0 && sliceLastPoint>0 && slicePoint==sliceLastPoint){
                      MaptoSubCategory(categoryCode,categoryName,rowIndex,rowLen,worksheet)
                  }
                  }
              }
      
              //fill the empty categorycode ane categoryname
              for(let rowIndex=1;rowIndex<=rowLen;rowIndex++){
                  let categoryCode=null
                  let categoryName=null
                  let row=worksheet.getRow(rowIndex)
                  if(row.getCell(1).value!=null){
                  categoryCode=row.getCell(1).value.toString().trim()
                  row.getCell(1).value=categoryCode
                  } else{
                  categoryCode=null
                  }
                  if(row.getCell(2).value!=null){
                  categoryName=row.getCell(2).value.toString().trim()
                  row.getCell(2).value=categoryName
                  } else{
                  categoryName=null
                  }
                  if(categoryCode==null && categoryName==null){
                  let preRow=worksheet.getRow(rowIndex-1)
                  row.getCell(1).value=preRow.getCell(1).value
                  row.getCell(2).value=preRow.getCell(2).value
                  }
              }
      
              //delete first level category
              for(let rowIndex=1;rowIndex<=rowLen-1;rowIndex++){
                  let categoryCode=null
                  let categoryName=null
                  let row=worksheet.getRow(rowIndex)
                  if(row.getCell(1).value!=null){
                  categoryCode=row.getCell(1).value.toString().trim()
                  row.getCell(1).value=categoryCode
                  } else{
                  categoryCode=null
                  }
                  if(row.getCell(2).value!=null){
                  categoryName=row.getCell(2).value.toString().trim()
                  row.getCell(2).value=categoryName
                  } else{
                  categoryName=null
                  }
                  if(categoryCode!=null){
                  let slicePoint=categoryCode.indexOf('.')
                  let thirdCell=row.getCell(3).value
                  let forthCell=row.getCell(4).value
                  if(categoryCode.length==4 && slicePoint<0&&thirdCell==null && forthCell==null){
                      let nextRow=worksheet.getRow(rowIndex+1)
                      if(hasSubCategory(categoryCode,nextRow)){
                      summarySheet.addRow(row.values)
                      worksheet.spliceRows(rowIndex,1)
                      }
                  }
                  }
              }
      
              //delete second level 123.01
              rowLen=worksheet.rowCount
              for(let rowIndex=1;rowIndex<=rowLen-1;rowIndex++){
                  let categoryCode=null
                  let categoryName=null
                  let row=worksheet.getRow(rowIndex)
                  if(row.getCell(1).value!=null){
                  categoryCode=row.getCell(1).value.toString().trim()
                  row.getCell(1).value=categoryCode
                  } else{
                  categoryCode=null
                  }
                  if(row.getCell(2).value!=null){
                  categoryName=row.getCell(2).value.toString().trim()
                  row.getCell(2).value=categoryName
                  } else{
                  categoryName=null
                  }
                  if(categoryCode!=null){
                  let slicePoint=categoryCode.indexOf('.')
                  let sliceLastPoint=categoryCode.lastIndexOf('.')
                  let thirdCell=row.getCell(3).value
                  let forthCell=row.getCell(4).value
                  if(slicePoint>0 && sliceLastPoint>0 && slicePoint==sliceLastPoint && thirdCell==null && forthCell==null){
                      let nextRow=worksheet.getRow(rowIndex+1)
                      if(hasSubCategory(categoryCode,nextRow)){
                      subSummarySheet.addRow(row.values)
                      worksheet.spliceRows(rowIndex,1)
                      }
                  }
                  }
              }
              
              // regular expression check the categoryCode
              rowLen=worksheet.rowCount
              for(let rowIndex=1;rowIndex<=rowLen;rowIndex++){
                  let dimCode=null
                  let dimName=null
                  let row=worksheet.getRow(rowIndex)
                  let categroryCode=row.getCell(1).value
                  if(row.getCell(3).value!=null){
                  dimCode=row.getCell(3).value.toString().trim()
                  }
      
                  if(row.getCell(4).value!=null){
                  dimName=row.getCell(4).value.toString().trim()
                  }
      
                  if(dimCode!=null && dimName!=null){
                  let checkBackValue=regularCheckDim(categroryCode)
                  let newColumn
                  switch(checkBackValue){
                      case 'cbcc':
                      newColumn=dimName.split('/')
                      row.getCell(13).value=newColumn[0]
                      row.getCell(14).value=newColumn[1]
                      row.getCell(15).value=newColumn[2]
                      row.getCell(16).value=newColumn[3]
                      break
                      case 'cb':
                      newColumn=dimName.split('/')
                      row.getCell(13).value=newColumn[0]
                      row.getCell(14).value=newColumn[1]
                      break
                      case 'cost':
                      row.getCell(13).value=dimName
                      break
                      default:
                      row.getCell(16).value=dimName
                  }
                  }
              }
      
              let targetSheet=targetWorkbook.addWorksheet("TBResult")
              rowLen=worksheet.rowCount
              for(let rowIndex=1;rowIndex<=rowLen;rowIndex++){
                  let row=worksheet.getRow(rowIndex)
                  let item11=row.getCell(5).value
                  let item12=row.getCell(6).value
                  let item2=row.getCell(7).value
                  let item3=row.getCell(8).value
                  let item41=row.getCell(11).value
                  let item42=row.getCell(12).value
      
                  if(item11!=null||item12!=null){
                  let targetRow=targetSheet.addRow()
                  targetRow.getCell(1).value="CNSH"
                  targetRow.getCell(2).value="TB"
                  targetRow.getCell(3).value=row.getCell(1).value
                  targetRow.getCell(4).value=row.getCell(2).value
                  targetRow.getCell(5).value=row.getCell(16).value
                  targetRow.getCell(6).value=monthValue
                  targetRow.getCell(7).value="CNY"
                  targetRow.getCell(9).value="CNY"
                  targetRow.getCell(12).value=row.getCell(14).value?row.getCell(14).value:"none"
                  targetRow.getCell(13).value=row.getCell(15).value?row.getCell(15).value:"none"
                  targetRow.getCell(14).value=row.getCell(13).value?row.getCell(13).value:"none"
      
                  let item1=item11?item11:item12
                  targetRow.getCell(8).value=item1
                  targetRow.getCell(10).value=item1
                  targetRow.getCell(11).value="1"
                  }
      
                  if(item2!=null){
                  let targetRow=targetSheet.addRow()
                  targetRow.getCell(1).value="CNSH"
                  targetRow.getCell(2).value="TB"
                  targetRow.getCell(3).value=row.getCell(1).value
                  targetRow.getCell(4).value=row.getCell(2).value
                  targetRow.getCell(5).value=row.getCell(16).value
                  targetRow.getCell(6).value=monthValue
                  targetRow.getCell(7).value="CNY"
                  targetRow.getCell(9).value="CNY"
                  targetRow.getCell(12).value=row.getCell(14).value?row.getCell(14).value:"none"
                  targetRow.getCell(13).value=row.getCell(15).value?row.getCell(15).value:"none"
                  targetRow.getCell(14).value=row.getCell(13).value?row.getCell(13).value:"none"
      
                  targetRow.getCell(8).value=item2
                  targetRow.getCell(10).value=item2
                  targetRow.getCell(11).value="2"
                  }
      
                  if(item3!=null){
                  let targetRow=targetSheet.addRow()
                  targetRow.getCell(1).value="CNSH"
                  targetRow.getCell(2).value="TB"
                  targetRow.getCell(3).value=row.getCell(1).value
                  targetRow.getCell(4).value=row.getCell(2).value
                  targetRow.getCell(5).value=row.getCell(16).value
                  targetRow.getCell(6).value=monthValue
                  targetRow.getCell(7).value="CNY"
                  targetRow.getCell(9).value="CNY"
                  targetRow.getCell(12).value=row.getCell(14).value?row.getCell(14).value:"none"
                  targetRow.getCell(13).value=row.getCell(15).value?row.getCell(15).value:"none"
                  targetRow.getCell(14).value=row.getCell(13).value?row.getCell(13).value:"none"
      
                  targetRow.getCell(8).value=item3
                  targetRow.getCell(10).value=item3
                  targetRow.getCell(11).value="3"
                  }
      
                  if(item41!=null||item42!=null){
                  let targetRow=targetSheet.addRow()
                  targetRow.getCell(1).value="CNSH"
                  targetRow.getCell(2).value="TB"
                  targetRow.getCell(3).value=row.getCell(1).value
                  targetRow.getCell(4).value=row.getCell(2).value
                  targetRow.getCell(5).value=row.getCell(16).value
                  targetRow.getCell(6).value=monthValue
                  targetRow.getCell(7).value="CNY"
                  targetRow.getCell(9).value="CNY"
                  targetRow.getCell(12).value=row.getCell(14).value?row.getCell(14).value:"none"
                  targetRow.getCell(13).value=row.getCell(15).value?row.getCell(15).value:"none"
                  targetRow.getCell(14).value=row.getCell(13).value?row.getCell(13).value:"none"
      
                  let item4=item41?item41:item42
                  targetRow.getCell(8).value=item4
                  targetRow.getCell(10).value=item4
                  targetRow.getCell(11).value="4"
                  }
              }
            } else{
                let targetWrongSheet=targetWorkbook.addWorksheet("Error")
                targetWrongSheet.getCell("A1").value="Bad data sourcing excel"
            }
            // saveToTargetFile(workbook,dirPath,"intermediate")
            saveToTargetFile(targetWorkbook,targetFile)
            // return {success:true,target:targetFile}

        })
        return {success:true,target:targetFile}

    } else{
        console.log("empty file")
        return {success:false,target:null}
    }
}

function MaptoSubCategory(categoryCode,categoryName,currentRow,rowLen,worksheet){
  for(let i=currentRow+1;i<=rowLen;i++){
    let nextRow=worksheet.getRow(i)
    let nextCategoryCode=null
    let nextCategoryName=null
    if(nextRow.getCell(1).value!=null){
      nextCategoryCode=nextRow.getCell(1).value.toString().trim()
    }
    if(nextRow.getCell(2).value!=null){
      nextCategoryName=nextRow.getCell(2).value.toString().trim()
    }
    if(nextCategoryCode!=null && nextCategoryName!=null){
      let sliceLastPoint=nextCategoryCode.lastIndexOf('.')
      if(sliceLastPoint>0){
        let nextCategoryCodePre=nextCategoryCode.slice(0,sliceLastPoint)
        if(nextCategoryCodePre==categoryCode){
          nextRow.getCell(2).value=categoryName+"--"+nextCategoryName.toString().trim()
        }
      }
    }
  }
}
function hasSubCategory(categoryCode,nextRow){
  let nextCategoryCode=nextRow.getCell(1).value
  if(nextCategoryCode!=null){
    if(nextCategoryCode==categoryCode){
      return true
    } else{
      let sliceLastPoint=nextCategoryCode.lastIndexOf('.')
      if(sliceLastPoint>0){
        let nextCategoryCodePre=nextCategoryCode.slice(0,sliceLastPoint)
        if(nextCategoryCodePre==categoryCode){
          return true
        }
      }
    }
  }
  return false
}
function saveToTargetFile(workbook,targetFile){
//   workbook.xlsx.writeFile("./"+dir+"/"+filepre+Date.now().toString()+".xlsx").then(function(){

//   })
     workbook.xlsx.writeFile(targetFile).then(function(){
        // console.log("success save to target")
    })
}
      //add three column 
      // 1)  6501.01.XX，对应维度为成本中心，品牌，渠道，客户
      // 2)  6502.01，对应维度为成本中心，品牌，渠道，客户
      // 3)  6503.XX，对应维度为成本中心，品牌，渠道，客户
      // 4)  6504.XX，对应维度为成本中心，品牌
      // 5)  6507.XX.XX，对应维度为成本中心
      // 6)  6601.XX，对应维度为成本中心
function regularCheckDim(dim){
  var cbcc=new RegExp("^6501.[0-9]*.?[0-9]*$|^6502.?[0-9]*$|^6503.?[0-9]*$")
  var cb=new RegExp("^6504")
  var cost=new RegExp("^6507.?[0-9]*$|^6601.?[0-9]*$")
  if(cbcc.test(dim)){
    return "cbcc"
  }
  if(cb.test(dim)){
    return "cb"
  }
  if(cost.test(dim)){
    return "cost"
  }
  return "default"
}
// readWorkbook()

