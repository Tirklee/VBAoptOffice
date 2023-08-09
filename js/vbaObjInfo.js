const xpath = require('xpath');
const {DOMParser} = require('xmldom');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

function getVBAObjInfo(filePath){
  //读取文件目录
  let fileList =  fs.readdirSync(filePath);
  let objStr = '<?xml version="1.0" encoding="UTF-8"?>\r\n';
  objStr+="<VbaObjs>\r\n";
  for(let file of fileList){
    if(file.endsWith(".html")){
      let obaPath=path.join(filePath,file);
      const fileData = fs.readFileSync(obaPath,"utf8");
      let doc = new DOMParser ().parseFromString(fileData);
      let headersTitle = xpath.select("//*[@id=\"nsrTitle\"]/b",doc);
      let objName = file.replace(".html","").replace(/\d/g,"");
      objStr+="\t<!--=========================="+objName+"===start==========================-->\r\n";
      objStr+="\t<!--"+headersTitle[0].textContent+"-->\r\n";
      let headersDesc = xpath.select("/html/body/div[2]/div/p[1]",doc);
      objStr+="\t<!--"+headersDesc[0].textContent+"-->\r\n";
      objStr+="\t<VbaObj id='"+objName.replace("对象","").trim()+"' name='"+headersTitle[0].textContent.trim()+"' desc='"+headersDesc[0].textContent+"'>\r\n";
      let methodList = xpath.select("//*[@id=\"vstable\"]/table",doc); 
      let i=0;
      for(let node of methodList){
        let childRows = node.childNodes;
        if(i==0){
          objStr+="\t\t<!--方法-->\r\n";  
          objStr+="\t\t<Methods>\r\n";
          for(let r=1;r<childRows.length;r++){
            let childColumns = childRows[r].childNodes;
            objStr+="\t\t\t<!-- methodName:名称 desc：说明-->\r\n";
            objStr+="\t\t\t<Item methodName=\'"+childColumns[1].textContent.trim().replaceAll(/\r?\n/g,"")+"\' desc=\'"+childColumns[2].textContent.trim().replaceAll(/\r?\n/g,"")+"\'/>\r\n";
          }
          objStr+="\t\t</Methods>\r\n" ;
        }else if(i==1){
          objStr+="\t\t<!--属性-->\r\n" ;
          objStr+="\t\t<AttrItems>\r\n";
          for(let r=1;r<childRows.length;r++){
            let childColumns = childRows[r].childNodes;
            objStr+="\t\t\t<!-- attrName:名称 attrType:数据类型(str字符串 list列表 obj对象)  desc：说明-->\r\n";
            objStr+="\t\t\t<Item attrName=\'"+childColumns[1].textContent.trim().replaceAll(/\r?\n/g,"")+"\' attrType='str' desc=\'"+childColumns[2].textContent.trim().replaceAll(/\r?\n/g,"")+"\'/>\r\n";
          }
          objStr+="\t\t</AttrItems>\r\n";
        }
        i++;
      }
      objStr+="\t</VbaObj>\r\n";
      objStr+="\t<!--=========================="+objName+"===end==========================-->\r\n";
    }
  }
  objStr+="</VbaObjs>\r\n";
  const saveObjXmlPath = path.join(filePath,"xml/VbaObjModel.xml");
  const writeStream = fs.createWriteStream(saveObjXmlPath);
  writeStream.write(objStr, () => {
    console.log('File written successfully.');
  });
  writeStream.end();
}
const filePath = "D:/tirklee/VBAOptWord/VBAoptOffice/js/vbaobjInfo";
getVBAObjInfo(filePath);
