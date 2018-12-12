/**
 *  2018/11/29  lize
 */
;(function(window,document){

    "use strict";
    
    var _self;

    var PoPmask;

    var loadingPop;

    var closeBtn;

    var labelDiv;

    var labelul;

    window.onload = function(){

        (function(){

            var loadingPop = document.createElement("div");

            loadingPop.setAttribute("class", "mask");

            loadingPop.setAttribute("id", "loadingMask");

            var TablePopHtml =
                "<div class = 'spinner'>"+
                "<div class = 'bounce1'></div>"+
                "<div class = 'bounce2'></div>"+
                "<div class = 'bounce3'></div>"+
                "</div>";

            loadingPop.innerHTML = TablePopHtml;

            document.body.appendChild(loadingPop);

        })();

        loadingPop = document.getElementById("loadingMask");

        // loadingPop.classList.add("displayBlock");
        //设置弹窗
        (function(){

            var TablePop = document.createElement("div");

            TablePop.setAttribute("class", "mask");

            TablePop.setAttribute("id", "mask");

            var TablePopHtml =
                "<div class = 'alertWarp'>"+
                "<h3><span id = 'titleSpan'>查看、修改所导入的数据</span><span class = 'claceSpan' id = 'closeBtn'>X</span></h3>" +
                "<div class = 'alert_content'>" +
                "<div class = 'table-cont' id = 'table-cont'>" +
                "<table cellspacing='0' cellpadding='0' border='0' class = 'tablePop' id = 'tablePop' width='100%'>" +
                "<thead id = 'theadPop'></thead>" +
                "<tbody id = 'tbodyPop'></tbody>" +
                "</table>" +
                "</div>" +
                "<div id = 'labelDiv' class = 'labelDiv' style='overflow: hidden'>" +
                "<ul id = 'labelul' style='overflow: hidden'>" +
                "</ul>"+
                "</div>"+
                "</div>"+
                "<div id = 'PoPfooter' class='commit right'></div>"+
                "</div>";

            TablePop.innerHTML = TablePopHtml;

            document.body.appendChild(TablePop);

            var tableCont = document.querySelector('.table-cont');

            tableCont.onscroll = function(){

                var TableScrollTop = tableCont.scrollTop;

                document.querySelector(".table-cont thead").style.transform = 'translateY(' + TableScrollTop + 'px)'

            };

        }())

        PoPmask = document.getElementById("mask");

        closeBtn = document.getElementById("closeBtn");

        labelDiv = document.getElementById("labelDiv");

      labelul = document.getElementById("labelul");

    };

    function extend(o,n,override) {
        
        for(var key in n){
            
            if(n.hasOwnProperty(key) && (!o.hasOwnProperty(key) || override)){
                
                o[key]=n[key];
                
            }
            
        }
        
        return o;
        
    }
    
    function ExcelImport(opt){
        
        this._initial(opt);
        
    }

    ExcelImport.prototype = {
        //初始化
        _initial:function(opt){
    
            _self = this;
    
            var def = {
                
                el:"",
                
                text:"导入",
    
                ImportImgUrl:'',
                
                succColor:'#000000',
    
                errColor:'red',
    
                ExcelRegulation:[],
                
            };
            
            this.def = extend(def,opt,true);

            this.setInnerHTML();

        },
        //设置按钮
        setInnerHTML:function(){

            if(this.def.el == ""){

                var importDiv = document.createElement('div');

                importDiv.id = "importDivId";

                document.body.appendChild(importDiv);

                this.container = document.getElementById('importDivId');

            }else{

                this.container = document.getElementById(this.def.el);

            }
            
            if(this.def.ImportImgUrl){
                
                this.content =
                        "<div style='width: 100%;height: 100%;'>" +
                            "<input type='file' style = 'display:none' id = 'upload' accept='.csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel'>" +
                            "<img style='width: 100%;height: 100%;cursor: pointer' id = 'Excelinport' src = "+this.def.ImportImgUrl+"/>" +
                        "</div>"
                
            }else{
    
                this.content =
                        "<div style='width: 100%;height: 100%;'>" +
                            "<input type='file' style = 'display:none' id = 'upload' accept='.csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel'>" +
                            "<button id = 'Excelinport' class = 'Excelinport' style='width: 100%;height: 100%'>"+this.def.text+"</button>" +
                        "</div>"
    
            }
    
            this.container.innerHTML = this.content;

            this.bingEvent();
            
        },
        //事件绑定
        bingEvent:function(){

            var ExcelinportLabel = document.getElementById("Excelinport");

            var inpuylabel = document.getElementById("upload");

            ExcelinportLabel.removeEventListener('click',function(){});

            ExcelinportLabel.addEventListener("click",function(){

                return inpuylabel.click();

            })
    
            inpuylabel.removeEventListener('change',function(){});

            inpuylabel.addEventListener('change',function(e){

                loadingPop.classList.remove("displayNone");

                loadingPop.classList.add("displayBlock");

                _self.excelAnalysis(e);

            });
            
        },
        //解析Excel
        excelAnalysis:function(e){
            
            var wb;
            
            var rABS;

            var persons = []; // 存储获取到的数据

            var SheetsAry = []; // 存储获取到的标签

            var workbook = null; //二进制的表格内容

            var fromTo = ""; //表格范围，可用于判断表头是否数量是否正确
        
            var fileData = e.target.files;
            
            if(!fileData){
                
                return ;
                
            }
    
            var f = fileData[0];

            _self.def.fileName = f.name

            _self.def.fileSuffix = _self.def.fileName.substring(_self.def.fileName.lastIndexOf(".")+1,_self.def.fileName.length);

            _self.def.fileSize = f.size

            _self.def.fileType = f.type

            var fileReader = new FileReader();
    
            //读取开始时
            fileReader.onloadstart = function(e){
        
                console.log('开始读取文件……');
        
            };
    
            //读取成功时
            fileReader.onload = function(e) {
    
                console.log("读取完毕，解析中……");

                try{

                    var resultData = e.target.result;

                    if (rABS) {

                      workbook = XLSX.read(btoa(this.fixdata(resultData)), {

                            // 手动转化
                            type: "base64"

                        });

                    } else {

                      workbook = XLSX.read(resultData, {

                            type: "binary"

                        });

                    }

                }
                catch (er) {

                    loadingPop.classList.remove("displayBlock");

                    loadingPop.classList.add("displayNone");

                    this.def.error("文件类型不正确,读取失败！");

                }

              // 遍历每张表读取
              for (var sheet in workbook.Sheets) {

                var sheetAry = [];

                if (workbook.Sheets.hasOwnProperty(sheet)) {

                  fromTo = workbook.Sheets[sheet]['!ref'];

                  console.log(fromTo);

                  if(fromTo!=undefined){

                    sheetAry = XLSX.utils.sheet_to_json(workbook.Sheets[sheet]);

                  }

                  // break; // 如果只取第一张表，就取消注释这行

                }

                if(sheetAry.length>0){

                  persons.push(sheetAry)

                  SheetsAry.push({sheet:sheet})

                }

              }

                // var Sucjson = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

                _self.primitiveExcelData = persons

                _self.SheetsAry = SheetsAry

                _self.dealFile(_self.primitiveExcelData,SheetsAry); // analyzeData: 解析导入数据
                
            };
    
            //读取失败时
            fileReader.onerror = function(er){
        
                console.log("文件读取失败……");

                loadingPop.classList.remove("displayBlock");

                loadingPop.classList.add("displayNone");

                this.def.error("表格中有错误数据");
                
            };
    
            if (this.rABS) {

                fileReader.readAsArrayBuffer(f);
        
            } else {
        
                fileReader.readAsBinaryString(f);
        
            }
            
        },
        //转换string
        fixdata:function(data) {
        
            // 文件流转BinaryString
            var o = "";
        
            var l = 0;
        
            var w = 10240;
        
            for (; l < data.byteLength / w; ++l) {
            
                o += String.fromCharCode.apply(
                    
                        null,
                    
                        new Uint8Array(data.slice(l * w, l * w + w))
            
                );
            
            }
        
            o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
        
            return o;
        
        },
        //处理解析后的xlsx
        dealFile:function(e,SheetsAry){

            loadingPop.classList.remove("displayBlock");

            loadingPop.classList.add("displayNone");

            if (e.length <= 0) {

                this.def.error("请导入正确信息!");
    
            }else{
            
                console.log("文件解析成功，渲染中……");

                this.def.Edata = JSON.parse(JSON.stringify(e));
                
                // this.def.ExcelData = this.assemblyData(this.def.Edata,this.def.ExcelRegulation,this.def.succColor,this.def.errColor);

                this.def.ExcelData = this.pluralSheerAssemblyData(this.def.Edata,this.def.ExcelRegulation,this.def.succColor,this.def.errColor);

                if(this.def.ExcelData == "000000"){

                    this.def.error("表结构错误！第一行不能有中文字符！");

                    return

                }

                if(this.def.ExcelData.length<=0){

                  this.def.error("数据处理错误,请联系管理员！");

                  return

                }
                
                this.Poprender(this.def.ExcelData,SheetsAry);

                console.log("文件渲染成功！")
            
            }
        
        },
        //处理数据
        assemblyData:function(dataAry,regulationAry,succColor,errColor){
        
            var list = dataAry;
        
            var newAry = [];
        
            var newhead = list.splice(0,1);
        
            var hObj = {};
        
            //处理头部
            for(var i in newhead[0]){

                if(newhead[0][i].indexOf("（") >-1){

                    var bef = newhead[0][i].substring(0,newhead[0][i].indexOf("（"));

                    var aft = newhead[0][i].substring(newhead[0][i].indexOf("（"),newhead[0][i].length);

                    newhead[0][i] = "<font color = 'red'>"+bef + "</font><br/>"+aft;

                }else if(newhead[0][i].indexOf("(") >-1){

                    var bef = newhead[0][i].substring(0,newhead[0][i].indexOf("("));

                    var aft = newhead[0][i].substring(newhead[0][i].indexOf("("),newhead[0][i].length);

                    newhead[0][i] = "<font color = 'red'>"+bef + "</font><br/>"+aft;

                }

                if(!Regular.isChinese(i)){
                
                    console.log(i);
                
                    return "000000"
                
                }
            
                hObj[i] = i;
            
            }
        
            newhead.unshift(hObj)
        
            var head = newhead[0];
        
            var H = [];
        
            for( var i in head){
            
                H.push(i);
            
            }
        
            for( var i = 0; i<list.length; i++){
            
                var index = i;
            
                var Ary = [];
            
                var c = 0;
            
                //处理表格中的空值；
                for( var j = 0; j<H.length; j++){
                
                    if(list[i][H[j]] == undefined){
                    
                        var t = JSON.stringify(list[i]).slice(1);
                    
                        var g = t.substring(0,t.length-1);
                    
                        var hj = g.split(',');
                    
                        var r = '"'+H[j]+'":'+ null;
                    
                        hj.splice(j,0,r);
                    
                        hj[0] = "{" + hj[0];
                    
                        hj[hj.length-1] = hj[hj.length-1] + "}"
                    
                        list[i] = JSON.parse(hj.join(","));
                    
                    }
                
                }
            
                //  变成二维数组并添加属性
                for(var km in list[i]){
                
                    var obj = {
                    
                        name:list[i][km],
                    
                        key: km,
                    
                        row:index,
                    
                        col:c,
                    
                        color:succColor,
                    
                        flag:true,
                    
                        innerText:list[i][km],
                    
                    };
                
                    var regObj = {
                    
                        name:list[i][km],
                    
                        col:obj.col,
                    
                        row:obj.row,
                    
                    };

                    if(list[i][km]!=null){

                        if(typeof list[i][km] == "number" ){

                            obj.name = Number(String(list[i][km]).replace(/\s+/g,""));

                            obj.innerText = Number(String(list[i][km]).replace(/\s+/g,""));

                            regObj.name = Number(String(list[i][km]).replace(/\s+/g,""));

                        }else{

                            obj.name = list[i][km].replace(/\s+/g,"");

                            obj.innerText = list[i][km].replace(/\s+/g,"");

                            regObj.name = list[i][km].replace(/\s+/g,"");

                        }

                    }
                
                    var flag = this.verifier(regulationAry,regObj)
                
                    if(!flag){
                    
                        obj.color = errColor;
                    
                        obj.flag = false;
                    
                    }
                
                    Ary.push(obj);
                
                    c++;
                
                }
            
                newAry.push(Ary);
            
            }
        
            var obj = {
            
                tHead:newhead,
            
                tBody:newAry
            
            };

            return obj;
        
        },
        //多个sheet
        pluralSheerAssemblyData(dataAry,regulationAry,succColor,errColor){

          var Sheet_tag = dataAry;

          var RERURNARY = [];

          for(var sheet = 0; sheet<Sheet_tag.length;sheet++){

            var list = Sheet_tag[sheet];

            var newAry = [];

            var newhead = list.splice(0,1);

            var hObj = {};

            //处理头部
            for(var i in newhead[0]){

              if(newhead[0][i].indexOf("（") >-1){

                var bef = newhead[0][i].substring(0,newhead[0][i].indexOf("（"));

                var aft = newhead[0][i].substring(newhead[0][i].indexOf("（"),newhead[0][i].length);

                newhead[0][i] = "<font color = 'red'>"+bef + "</font><br/>"+aft;

              }else if(newhead[0][i].indexOf("(") >-1){

                var bef = newhead[0][i].substring(0,newhead[0][i].indexOf("("));

                var aft = newhead[0][i].substring(newhead[0][i].indexOf("("),newhead[0][i].length);

                newhead[0][i] = "<font color = 'red'>"+bef + "</font><br/>"+aft;

              }

              if(!Regular.isChinese(i)){

                console.log(i);

                return "000000"

              }

              hObj[i] = i;

            }

            newhead.unshift(hObj)

            var head = newhead[0];

            var H = [];

            for( var i in head){

              H.push(i);

            }

            for( var i = 0; i<list.length; i++){

              var index = i;

              var Ary = [];

              var c = 0;

              //处理表格中的空值；
              for( var j = 0; j<H.length; j++){

                if(list[i][H[j]] == undefined){

                  var t = JSON.stringify(list[i]).slice(1);

                  var g = t.substring(0,t.length-1)

                  var hj = g.split(',')

                  var r = '"'+H[j]+'":'+ null;

                  hj.splice(j,0,r);

                  hj[0] = "{" + hj[0];

                  hj[hj.length-1] = hj[hj.length-1] + "}"

                  list[i] = JSON.parse(hj.join(","));

                }

              }

              //  变成二维数组并添加属性
              for(var km in list[i]){

                var obj = {

                  name:list[i][km],

                  key: km,

                  row:index,

                  col:c,

                  sheet:Number(sheet)+1,

                  color:succColor,

                  flag:true,

                  innerText:list[i][km],

                };

                var regObj = {

                  name:list[i][km],

                  col:obj.col,

                  sheet:Number(sheet)+1,

                  row:obj.row,

                };

                if(list[i][km]!=null){

                  if(typeof list[i][km] == "number"){

                    obj.name = Number(String(list[i][km]).replace(/\s+/g,""));

                    obj.innerText = Number(String(list[i][km]).replace(/\s+/g,""));

                    regObj.name = Number(String(list[i][km]).replace(/\s+/g,""));

                  }else{

                    obj.name = list[i][km].replace(/\s+/g,"");

                    obj.innerText = list[i][km].replace(/\s+/g,"");

                    regObj.name = list[i][km].replace(/\s+/g,"");

                  }

                }

                var flag = this.verifier(regulationAry,regObj)

                if(!flag){

                  obj.color = errColor;

                  obj.flag = false;

                }

                Ary.push(obj);

                c++;

              }

              newAry.push(Ary);

            }

            var obj = {

              tHead:newhead,

              tBody:newAry

            }

            RERURNARY.push(obj);

          }

          return RERURNARY;

        },
        //验证函数
        verifier:function(regulationAry,regobj){
        
            var name = regobj.name;

            var col = regobj.col;
        
            var row = regobj.row;

            var sheet = 'sheet'+regobj.sheet;

            var reg = [];
        
            var flag = true;
        
            if(regulationAry!=undefined && regulationAry!=null){
            
                for(var i = 0; i< regulationAry.length; i++ ){

                  if(sheet == regulationAry[i].sheet){

                    for(var j = 0; j < regulationAry[i].info.length; j++){

                      if(regulationAry[i].info[j].index == col){

                        reg = regulationAry[i].info[j].reg;

                        if(name!==null || name!=undefined){

                          for( var k = 0; k<reg.length; k++){

                            flag = Regular[reg[k].name](name);

                          }

                        }else{

                          flag = false;

                        }

                      }

                    }

                  }
                
                }
            
            }
        
            return flag;
        
        },
        //渲染弹窗及事件
        Poprender:function(arr,SheetsAry){

          _self = this;

          var labelLiHTML = '';

          var theadtr = "";

          var tbodytr = "";

          var footerBtn = "";

          //渲染sheet标签

          if(SheetsAry.length>1){

            for(var sheet = 0; sheet<SheetsAry.length;sheet++){

              if(sheet ==0){

                labelLiHTML +="<li data-sheet = '"+(Number(sheet))+"' class = 'active'>"+SheetsAry[sheet].sheet+"</li>"

              }else{

                labelLiHTML +="<li data-sheet = '"+(Number(sheet))+"'>"+SheetsAry[sheet].sheet+"</li>"

              }

            }

          }
          
          labelul.innerHTML = "";

          labelul.innerHTML = labelLiHTML;

          for(var i = 0;i<arr[0].tHead.length;i++){

              if(i == 1){

                  theadtr+="<tr>";

                  for(var j in arr[0].tHead[i] ){

                      theadtr += "<th>"+arr[0].tHead[i][j]+"</th>"

                  }

                  theadtr +="</tr>"

              }

          }

          for(var i = 0; i <arr[0].tBody.length;i++){

              tbodytr += "<tr>";

              for(var j in arr[0].tBody[i]){

                  if(arr[0].tBody[i][j].name != null){

                      tbodytr +=

                          "<td contentEditable = 'plaintext-only' class = 'tobytd' title = '点击修改' data-sheet = '"+arr[0].tBody[i][j].sheet+"'  data-col = '"+arr[0].tBody[i][j].col+"' data-row = '"+arr[0].tBody[i][j].row+"' data-key = '"+arr[0].tBody[i][j].key+"' style = 'color:"+arr[0].tBody[i][j].color+"'>"+arr[0].tBody[i][j].name+"</td>"

                  }else{

                      tbodytr +=

                          "<td contentEditable = 'plaintext-only' class = 'tobytd' title = '点击修改' data-sheet = '"+arr[0].tBody[i][j].sheet+"'  data-col = '"+arr[0].tBody[i][j].col+"' data-row = '"+arr[0].tBody[i][j].row+"' data-key = '"+arr[0].tBody[i][j].key+"' style = 'color:"+arr[0].tBody[i][j].color+"'></td>"

                  }

              }

              tbodytr += "</tr>"

          }

          footerBtn =

              "<a href='javascript:;' id = 'ImportPopsave' class='save commit_default'>确定</a>"+
              "<a href='javascript:;' id = 'ImportPopcancel' class='cancel commit_default'>取消</a>";
            
            var theadPop = document.getElementById("theadPop");
            
            var tbodyPop = document.getElementById("tbodyPop");

            var footerPopNode = document.getElementById("PoPfooter");

            theadPop.innerHTML = "";

            tbodyPop.innerHTML = "";

            footerPopNode.innerHTML = "";

            theadPop.innerHTML = theadtr;

            tbodyPop.innerHTML = tbodytr;

            footerPopNode.innerHTML = footerBtn

            var PopSave = document.getElementById("ImportPopsave");
            
            var Popcancel = document.getElementById("ImportPopcancel");

            PoPmask.classList.remove("displayNone");

            PoPmask.classList.add("displayBlock");

            //td失去焦点
            tbodyPop.addEventListener('focusout', function(e){

                var targetSheet = e.target.getAttribute("data-sheet")

                var data = _self.def.ExcelData[(Number(targetSheet)-1)].tBody;

                var targetHtml;

                if(e.target.innerHTML!=null){

                    targetHtml = e.target.innerHTML.replace(/\s+/g,"");

                    e.target.innerHTML = e.target.innerHTML.replace(/\s+/g,"");

                }else{

                    targetHtml = e.target.innerHTML;

                    e.target.innerHTML = e.target.innerHTML;

                }

                var targetRow = e.target.getAttribute("data-row");

                var targetCol = e.target.getAttribute("data-col");

                var obj = {

                    name:targetHtml,

                    col:targetCol,

                    row:targetRow,

                    sheet:targetSheet

                }

                var flag = _self.verifier(_self.def.ExcelRegulation,obj);

                for(var i = 0; i<data.length;i++){

                    if(i == targetRow){

                        for(var j = 0; j<data[i].length;j++){

                            if(j == targetCol){

                                data[i][j].name = targetHtml;

                                data[i][j].innerText = targetHtml;

                                if(flag){

                                    e.target.style.color = _self.def.succColor;

                                    data[i][j].color = _self.def.succColor;

                                    data[i][j].flag = true;

                                }else{

                                    e.target.style.color = _self.def.errColor;

                                    data[i][j].color = _self.def.errColor;

                                    data[i][j].flag = false;

                                }

                            }

                        }

                    }

                };

            }, false)

            //点击sheet标签
            labelul.removeEventListener("click",function(){});

            labelul.addEventListener("click",function(e){

              if(e.target.nodeName == "LI"){

                theadPop.innerHTML = "";

                tbodyPop.innerHTML = "";

                theadtr = '';

                tbodytr = '';

                var sheetIndex = e.target.getAttribute("data-sheet")

                var SiblingNode = e.target.parentElement.childNodes;

                if(SiblingNode){

                  for(var i = 0; i < SiblingNode.length; i++){

                    SiblingNode[i].classList.remove("active");

                  }

                }

                e.target.classList.add("active");

                for(var i = 0;i<arr[sheetIndex].tHead.length;i++){

                  if(i == 1){

                    theadtr+="<tr>";

                    for(var j in arr[sheetIndex].tHead[i] ){

                      theadtr += "<th>"+arr[sheetIndex].tHead[i][j]+"</th>"

                    }

                    theadtr +="</tr>"

                  }

                }

                for(var i = 0; i <arr[sheetIndex].tBody.length;i++){

                  tbodytr += "<tr>";

                  for(var j in arr[sheetIndex].tBody[i]){

                    if(arr[sheetIndex].tBody[i][j].name != null){

                      tbodytr +=

                        "<td contentEditable = 'plaintext-only' class = 'tobytd' title = '点击修改' data-sheet = '"+arr[sheetIndex].tBody[i][j].sheet+"'  data-col = '"+arr[sheetIndex].tBody[i][j].col+"' data-row = '"+arr[sheetIndex].tBody[i][j].row+"' data-key = '"+arr[sheetIndex].tBody[i][j].key+"' style = 'color:"+arr[sheetIndex].tBody[i][j].color+"'>"+arr[sheetIndex].tBody[i][j].name+"</td>"

                    }else{

                      tbodytr +=

                        "<td contentEditable = 'plaintext-only' class = 'tobytd' title = '点击修改' data-sheet = '"+arr[sheetIndex].tBody[i][j].sheet+"'  data-col = '"+arr[sheetIndex].tBody[i][j].col+"' data-row = '"+arr[sheetIndex].tBody[i][j].row+"' data-key = '"+arr[sheetIndex].tBody[i][j].key+"' style = 'color:"+arr[sheetIndex].tBody[i][j].color+"'></td>"

                    }

                  }

                  tbodytr += "</tr>"

                }

                theadPop.innerHTML = theadtr;

                tbodyPop.innerHTML = tbodytr;

              }

            });

            //清除按钮
            Popcancel.removeEventListener("click",function(){});

            Popcancel.addEventListener("click",function(){

                PoPmask.classList.remove("displayBlock");

                PoPmask.classList.add("displayNone");

            });

            //关闭按钮
            closeBtn.removeEventListener("click",function(){});

            closeBtn.addEventListener("click",function(){

                PoPmask.classList.remove("displayBlock");

                PoPmask.classList.add("displayNone");

            });

            //保存按钮
            PopSave.removeEventListener("click",function(){});

            PopSave.addEventListener("click",function(){

                if(_self.comparisonData() == false){

                    PoPmask.classList.add("displayBlock");

                }else{

                    _self.def.success(_self.comparisonData());

                    PoPmask.classList.remove("displayBlock");

                    PoPmask.classList.add("displayNone");

                }

            });

        },
        //保存前比对数据
        comparisonData(){

            var newAry = [];

            var Ary = this.def.ExcelData;

            if(Ary && Ary.length>0){

              for(var sheet = 0; sheet<Ary.length; sheet++) {

                var formattingHeadAry = Ary[sheet].tHead;

                var headAry = JSON.parse(JSON.stringify(formattingHeadAry));

                headAry[1] = _self.primitiveExcelData[sheet][0]

                var bodyAry = Ary[sheet].tBody;

                var objData = {

                  headData: headAry,

                  formattingHeadAry:formattingHeadAry,

                  bodyData: [],

                  fileName: this.def.fileName,

                  fileSuffix: this.def.fileSuffix,

                  fileSize: this.def.fileSize,

                  fileType: this.def.fileType,

                };

                if (bodyAry.length > 0) {

                  for (var i = 0; i < bodyAry.length; i++) {

                    var obj = {};

                    for (var j = 0; j < bodyAry[i].length; j++) {

                      if (bodyAry[i][j].flag) {

                        obj[bodyAry[i][j].key] = bodyAry[i][j].innerText;

                      } else {

                        this.def.error("表格中有错误数据!")

                        return false;

                      }

                    }

                    objData.bodyData.push(obj);

                  }

                }

                newAry.push(objData);

              }

            }

           return newAry;

        },
        
    };
    
    window.ExcelImport = ExcelImport;

    function ExcelExport(opt){

        this._initial(opt);

    }

    ExcelExport.prototype = {

        _initial:function(opt){

            var def = {

                el:'',

                color:"#000000",

                ExportImgUrl:'',

                text:'导出',

                ExcelExportData:[],

                ExportType:'xlsx',

                fileName:'下载.xlsx',

                SheetsAry:[]

            }

            this.def = extend(def,opt,true);

            this.setInnerHTML();

        },
        //设置按钮
        setInnerHTML:function(){

            if(this.def.el == ""){

                var exportDiv = document.createElement('div');

                exportDiv.id = "exportDivId";

                document.body.appendChild(exportDiv);

                this.container = document.getElementById('exportDivId');

            }else{

                this.container = document.getElementById(this.def.el);

            }

            if(this.def.ExportImgUrl){

                this.content =
                    "<div style='width: 100%;height: 100%;'>" +
                    "<input type='file' style = 'display:none' id = 'upload'>" +
                    "<img style='width: 100%;height: 100%;' src = "+this.def.ExportImgUrl+"/>" +
                    "</div>"

            }else{

                this.content =
                    "<div style='width: 100%;height: 100%;'>" +
                    "<input type='file' style = 'display:none' id = 'upload'>" +
                    "<button id = 'ExcelExport' class = 'ExcelExport' style='width: 100%;height: 100%'>"+this.def.text+"</button>" +
                    "</div>"

            }

            this.container.innerHTML = this.content;

        },
        //外暴露方法
        exportExcel:function(ArrayData){

          var flag = false;

          if(ArrayData && ArrayData.length>0) {

            for (var sheet = 0; sheet < ArrayData.length; sheet++) {

              if (ArrayData[sheet].headData && ArrayData[sheet].headData.length > 0) {

                if (ArrayData[sheet].bodyData && ArrayData[sheet].bodyData.length > 0) {

                  flag = true;

                }else{

                  flag = false;

                  return

                }

              }else{

                flag = false;

                return

              }

              if(ArrayData[sheet].fileSuffix){

                this.def.fileSuffix = ArrayData[sheet].fileSuffix;

              }

              if(ArrayData[sheet].fileSize){

                this.def.fileSize = ArrayData[sheet].fileSize;

              }

              if(ArrayData[sheet].fileType){

                this.def.fileType =  ArrayData[sheet].fileType

              }

            }

          }else{

            flag = false;

            return

          }

          if(flag){

            this.def.ExportData = ArrayData;

            this.def.ExcelData = this.disposeData(this.def.ExportData);

            if(this.def.ExcelData!==null){

              this.Poprender(this.def.ExcelData,this.def.SheetsAry);

            }else{

              this.def.error('数据格式错误,解析失败,请检查Excel数据')

            }

          }else{

            this.def.error('数据格式错误,解析失败,请检查Excel数据')

            return

          }

        },
        // 处理显示数据
        disposeData:function(ExPortArray){

          var RETURNARRAY = [];

          var awaitArray = JSON.parse(JSON.stringify(ExPortArray))

          if(awaitArray){

            for(var sheet = 0; sheet < awaitArray.length;sheet++ ){

              var headAry = awaitArray[sheet].headData;

              var Ary = awaitArray[sheet].bodyData;

              var formattingHeadAry = awaitArray[sheet].formattingHeadAry

              var newAry = [];

              for(var i = 0; i<Ary.length; i++){

                var index = i;

                var arr = [];

                var c = 0;

                for (var j in Ary[i]){

                  var obj = {

                    name:Ary[i][j],

                    row:index,

                    col:c

                  }

                  arr.push(obj);

                  c++

                }

                newAry.push(arr);

              }

              var obj = {

                tHead:headAry,

                tBody:newAry,

                formattingHeadAry:formattingHeadAry

              }

              RETURNARRAY.push(obj);

              this.def.SheetsAry.push({sheet:'Sheet'+(Number(sheet)+1)});

            }

          }

            return RETURNARRAY;

        },
        //渲染弹窗及事件
        Poprender:function(arr,SheetsAry){

            _self = this;

            var theadtr = "";

            var tbodytr = "";

            var footerBtn = ""

            var labelLiHTML = '';

            //渲染sheet标签

            if(SheetsAry.length>1){

              for(var sheet = 0; sheet<SheetsAry.length;sheet++){

                if(sheet ==0){

                  labelLiHTML +="<li data-sheet = '"+(Number(sheet))+"' class = 'active'>"+SheetsAry[sheet].sheet+"</li>"

                }else{

                  labelLiHTML +="<li data-sheet = '"+(Number(sheet))+"'>"+SheetsAry[sheet].sheet+"</li>"

                }

              }

            }

            labelul.innerHTML = "";

            labelul.innerHTML = labelLiHTML;

            for(var i = 0;i<arr[0].tHead.length;i++){

                if(i == 1){

                    theadtr+="<tr>";

                    for(var j in arr[0].tHead[i] ){

                        theadtr += "<th>"+arr[0].tHead[i][j]+"</th>"

                    }

                    theadtr +="</tr>"

                }

            }

            for(var i = 0; i <arr[0].tBody.length;i++){

                tbodytr += "<tr>";

                for(var j = 0; j <arr[0].tBody[i].length; j++){

                    if(arr[0].tBody[i][j].name != null){

                        tbodytr +=

                            "<td class = 'tobytd' title = ' 查看' style = 'color:"+_self.def.color+"'>"+arr[0].tBody[i][j].name+"</td>"

                    }else{

                        tbodytr +=

                            "<td class = 'tobytd' title = ' 查看' style = 'color:"+_self.def.color+"'></td>"

                    }

                }

                tbodytr += "</tr>"

            }

            footerBtn =

                "<a href='javascript:;' id = 'ExportPopsave' class='save commit_default'>确定</a>"+
                "<a href='javascript:;' id = 'ExportPopcancel' class='cancel commit_default'>清空</a>";


            var theadPop = document.getElementById("theadPop");

            var tbodyPop = document.getElementById("tbodyPop");

            var footerPopNode = document.getElementById("PoPfooter");

            theadPop.innerHTML = "";

            tbodyPop.innerHTML = "";

            footerPopNode.innerHTML = "";

            theadPop.innerHTML = theadtr;

            tbodyPop.innerHTML = tbodytr;

            footerPopNode.innerHTML = footerBtn;

            var ExportPopSave = document.getElementById("ExportPopsave");

            var ExportPopcancel = document.getElementById("ExportPopcancel");

            PoPmask.classList.remove("displayNone");

            PoPmask.classList.add("displayBlock");

            //关闭按钮
            ExportPopcancel.removeEventListener("click",function(){});

            ExportPopcancel.addEventListener("click",function(){

                PoPmask.classList.remove("displayBlock");

                PoPmask.classList.add("displayNone");

            });
            //保存按钮
            ExportPopSave.removeEventListener("click",function(){});

            ExportPopSave.addEventListener("click",function(){

              var SheetAry = [];

              var newSheetAry = JSON.parse(JSON.stringify(_self.def.ExportData));

              for(var i = 0; i<newSheetAry.length;i++){

                newSheetAry[i].bodyData.unshift(newSheetAry[i].headData[1]);

                SheetAry.push(newSheetAry[i].bodyData)

              }

                _self.downloadMater(SheetAry);

                _self.def.success("true");

                PoPmask.classList.remove("displayBlock");

                PoPmask.classList.add("displayNone");

            });

          //点击sheet标签
          labelul.removeEventListener("click",function(){});

          labelul.addEventListener("click",function(e){

            if(e.target.nodeName == "LI"){

              theadPop.innerHTML = "";

              tbodyPop.innerHTML = "";

              theadtr = '';

              tbodytr = '';

              var sheetIndex = e.target.getAttribute("data-sheet")

              var SiblingNode = e.target.parentElement.childNodes;

              if(SiblingNode){

                for(var i = 0; i < SiblingNode.length; i++){

                  SiblingNode[i].classList.remove("active");

                }

              }

              e.target.classList.add("active");

              for(var i = 0;i<arr[sheetIndex].tHead.length;i++){

                if(i == 1){

                  theadtr+="<tr>";

                  for(var j in arr[sheetIndex].tHead[i] ){

                    theadtr += "<th>"+arr[sheetIndex].tHead[i][j]+"</th>"

                  }

                  theadtr +="</tr>"

                }

              }

              for(var i = 0; i <arr[sheetIndex].tBody.length;i++){

                tbodytr += "<tr>";

                for(var j in arr[sheetIndex].tBody[i]){

                  if(arr[sheetIndex].tBody[i][j].name != null){

                    tbodytr +=

                      "<td contentEditable = 'plaintext-only' class = 'tobytd' title = '点击修改' data-sheet = '"+arr[sheetIndex].tBody[i][j].sheet+"'  data-col = '"+arr[sheetIndex].tBody[i][j].col+"' data-row = '"+arr[sheetIndex].tBody[i][j].row+"' data-key = '"+arr[sheetIndex].tBody[i][j].key+"' style = 'color:"+arr[sheetIndex].tBody[i][j].color+"'>"+arr[sheetIndex].tBody[i][j].name+"</td>"

                  }else{

                    tbodytr +=

                      "<td contentEditable = 'plaintext-only' class = 'tobytd' title = '点击修改' data-sheet = '"+arr[sheetIndex].tBody[i][j].sheet+"'  data-col = '"+arr[sheetIndex].tBody[i][j].col+"' data-row = '"+arr[sheetIndex].tBody[i][j].row+"' data-key = '"+arr[sheetIndex].tBody[i][j].key+"' style = 'color:"+arr[sheetIndex].tBody[i][j].color+"'></td>"

                  }

                }

                tbodytr += "</tr>"

              }

              theadPop.innerHTML = theadtr;

              tbodyPop.innerHTML = tbodytr;

            }

          });

        },
        //创建下载链接
        saveAs:function(obj, fileName) {//当然可以自定义简单的下载文件实现方式

            var tmpa = document.createElement("a");

            tmpa.download = fileName || "下载";

            tmpa.href = URL.createObjectURL(obj); //绑定a标签

            tmpa.click(); //模拟点击实现下载

            setTimeout(function () { //延时释放

                URL.revokeObjectURL(obj); //用URL.revokeObjectURL()来释放这个object URL

            }, 100);

        },
        //导出xlsx
        downloadMater:function(Ary){

            // const defaultCellStyle =  { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}};

            // const wopts = { bookType:'xlsx', bookSST:false, type:'binary', defaultCellStyle: defaultCellStyle, showGridLines: false};

            try{

                // var wopts = { bookType:this.def.ExportType, bookSST:true, type:'binary'};
                //
                // var wb = { SheetNames: ['Sheet1'], Sheets: {}, Props: {} };
                //
                // wb.Sheets['Sheet1'] = XLSX.utils.json_to_sheet(data)

                var wopts = { bookType:this.def.ExportData[0].fileSuffix, bookSST:true, type:'binary'};

                var wb = { SheetNames: [], Sheets: {}, Props: {} };

              if(Ary){

                for(var i = 0 ;i < Ary.length; i++){

                  wb.SheetNames.push('Sheet'+(i+1));

                  wb.Sheets['Sheet'+(i+1)] = XLSX.utils.json_to_sheet(Ary[i])

                }

              }

                //创建二进制对象写入转换好的字节流
                var tmpDown =  new Blob([this.s2ab(XLSX.write(wb, wopts))], { type: "application/octet-stream" })

                this.saveAs(tmpDown,this.def.ExportData[0].fileName)

            }
            catch (er) {

                console.log(er);

                this.def.error('下载格式设置错误,无法进行下载，请联系管理员！')

            }


        },
        //字符串转字符流
        s2ab:function(s) {

            if (typeof ArrayBuffer !== 'undefined') {

                var buf = new ArrayBuffer(s.length);

                var view = new Uint8Array(buf);

                for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;

                return buf;

            } else {

                var buf = new Array(s.length);

                for (var i = 0; i != s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF;

                return buf;

            }

        },

    }

    window.ExcelExport= ExcelExport;

}(window,document));
