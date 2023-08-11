const xpath = require('xpath');
const {DOMParser} = require('xmldom');
const axios = require('axios');
const fs = require('fs');
const builder  = require('xmlbuilder'); 

// 定义多个API接口URL和请求参数
const apiList = [];
function getEnumList(){
  return axios.get("https://learn.microsoft.com/zh-cn/office/vba/api/word(enumerations)").then(response => {
    // 处理获取到的数据
    let data = response.data;
    let doc = new DOMParser ().parseFromString(data);
    let nodelistx = xpath.select('//main/div[3]/ul[1]/li', doc);
    nodelistx.forEach(itemx=>{
      let node =  itemx.textContent;
      let item = node.toLowerCase();
      let apiObj = {};
      apiObj.url = 'https://learn.microsoft.com/zh-cn/office/vba/api/word.'+item;
      apiObj.params ={};
      apiObj.EnumField=node;
      apiList.push(apiObj);
    });
  }).catch(error => {
    console.error(error);
  });
}
// 定义按顺序执行Ajax请求的函数
function fetchApiList(apiList) {
  const promises = apiList.map(api => {
    return axios.get(api.url, {
      params: api.params,
      headers: {
      }
    }).then(response => {
      let dataObj = {};
      dataObj["data"]=response.data;
      dataObj["api"]=api;
      return dataObj;
    }).catch(error => {
      console.error(error);
    });
  });
  return Promise.all(promises);
}

getEnumList().then(()=>{
  
  // 按顺序循环获取API接口的数据
  fetchApiList(apiList).then((nodelist) => {
    let wordEnum = builder.create("WordEnum", { version: "1.0", encoding: "UTF-8" })
    nodelist.forEach(node1=>{
        // 处理获取到的数据
        let data = node1.data;
        let api = node1.api;
        let doc = new DOMParser ().parseFromString(data);
        let enumEn = xpath.select('//main/div[3]/h1', doc).at(0);
        let enumCh = xpath.select('//main/div[3]/p[1]', doc).at(0);
        let comment= wordEnum.com(enumEn.textContent.replaceAll("\r\n","").trim()+"===="+enumCh.textContent.replaceAll("\r\n","").trim());
        let enumX =  wordEnum.ele("Enum");
        enumX.att("id",api.EnumField);  
        // 使用XPath查找所有a标签的href属性值
        let nodelist = xpath.select('//table/tbody/tr', doc);
        nodelist.forEach(node=>{
          let name = node.childNodes[1].textContent;
          let value = node.childNodes[3].textContent;
          let description = node.childNodes[5].textContent.replaceAll("\r\n","").trim();
          let item = enumX.ele("Item");
          item.att("id",name); 
          item.att("value",value); 
          item.att("description",description); 
        });
    });
    let xml = wordEnum.end({ pretty: true});
    console.log(xml);
    const filename = 'D:/tirklee/VBAOptWord/VBAoptOffice/js/vbaobjInfo/xml/WordEnum.xml';
    fs.writeFile(filename, xml.toString(), (err) => {  
      if (err) throw err;  
      console.log('XML has been written to '+filename+' file');  
    });
  }).catch(error => {
  console.error(error);
  });
});


