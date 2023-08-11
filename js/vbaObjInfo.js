const xpath = require("xpath");
const {DOMParser} = require("xmldom");
const axios = require("axios");
const fs = require("fs");
const path = require("path");
const builder  = require('xmlbuilder'); 

function getVBAObjInfo(filePath){
  //读取文件目录
  let fileList =  fs.readdirSync(filePath);
  let objList = [];
  for(let file of fileList){
    if(file.endsWith(".html")){
      let objName = file.replace(".html","").replace(/\d/g,"");
      let id = objName.replace("对象","").trim()
      objList.push(id);
    }
  }
  let vbaObjs = builder.create("VbaObjs", { version: "1.0", encoding: "UTF-8" })
  for(let file of fileList){
    if(file.endsWith(".html")){
      let obaPath=path.join(filePath,file);
      const fileData = fs.readFileSync(obaPath,"utf8");
      let doc = new DOMParser ().parseFromString(fileData);
      let headersTitle = xpath.select("//*[@id=\"nsrTitle\"]/b",doc);
      let objName = file.replaceAll(".html","").replaceAll(/\d/g,"");
      vbaObjs.com("=========================="+objName+"===start==========================");
      vbaObjs.com(headersTitle[0].textContent);
      let headersDesc = xpath.select("/html/body/div[2]/div/p[1]",doc);
      vbaObjs.com(headersDesc[0].textContent);
      let vbaObj = vbaObjs.ele("VbaObj"); 
      vbaObj.att("id",objName.replaceAll("对象","").trim());
      vbaObj.att("name",headersTitle[0].textContent);
      vbaObj.att("desc",headersDesc[0].textContent);
      let methodList = xpath.select("//*[@id=\"vstable\"]/table",doc); 
      let i=0;
      for(let node of methodList){
        let childRows = node.childNodes;
        if(i==0){
          vbaObj.com("方法");
          let methods = vbaObj.ele("Methods");
          methods.att("id","Methods");
          for(let r=1;r<childRows.length;r++){
            let childColumns = childRows[r].childNodes;
            methods.com("methodName:名称 desc：说明");
            let methodName = childColumns[1].textContent;
            let desc = childColumns[2].textContent.replaceAll("\r\n","").trim();
            let item = methods.ele("Item");
            item.att("id",methodName);
            item.att("methodName",methodName);
            item.att("desc",desc);
          }
        }else if(i==1){
          vbaObj.com("属性");
          let attrItems = vbaObj.ele("AttrItems");
          attrItems.att("id","AttrItems");
          for(let r=1;r<childRows.length;r++){
            let childColumns = childRows[r].childNodes;
            attrItems.com("attrName:名称 attrType:数据类型(str字符串 list列表 obj对象)  desc：说明");
            let desc = childColumns[2].textContent.replaceAll("\r\n","").trim();
            let attrName = childColumns[1].textContent;
            let attrType = "str";
            objList.forEach(objNameX => {
              if(desc.indexOf("对象")>0 && desc.indexOf(objNameX)>0){
                attrType = "obj"; 
              }else if(desc.indexOf("集合")>0 && desc.indexOf(objNameX)>0){
                attrType = "list";
              }
            });
            if(attrName=="Count"){
              attrType = "str"
            }
            let item = attrItems.ele("Item");
            item.att("id",attrName);
            item.att("attrName",attrName);
            item.att("attrType",attrType);
            item.att("desc",desc);
          }
        }
        i++;
      }
      vbaObjs.com("=========================="+objName+"===end==========================");
    }
  }
  let xml = vbaObjs.end({ pretty: true});
  console.log(xml);
  const saveObjXmlPath = path.join(filePath,"xml/VbaObjModel.xml");
  fs.writeFile(saveObjXmlPath, xml.toString(), (err) => {  
    if (err) throw err;  
    console.log('XML has been written to '+saveObjXmlPath+' file');  
  });
}
const filePath = "D:/tirklee/VBAOptWord/VBAoptOffice/js/vbaobjInfo";
getVBAObjInfo(filePath);
