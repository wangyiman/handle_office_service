const Koa = require('koa');
const fs = require('fs');
const mimes = require('./mimes')
const path = require('path')
const generateExcel = require('./utils/generateExcel');
const Excel = require('exceljs');

const app = new Koa();
const staticPath = './tempExcel';    

app.use(async (ctx, next)=>{
    //当请求时GET请求时，显示表单让用户填写
    if(ctx.url==='/' && ctx.method === 'GET'){
      var header = {
        firstLine: 'name',
        secondLine: 'age'
      }
      var data = [{
        'name': 'wangyiman',
        'age': '13'
      },{
        'name':'zhangsan',
        'age': '40'
      },{
        'name': 'zhangeryang',
        'age':'30'
      }];

      data.unshift(header);
      var finnalData = {
        data: data
      }
      var finalStr = JSON.stringify(finnalData);
        let html =`
            <h1>Koa2 request post demo</h1>
            <form method="POST"  action="/bomlist">
                <input name='webSite' value=${finalStr}/><br/>
                <button type="submit">submit</button>
            </form>
        `;
        ctx.response.type = 'text/html';        
        ctx.response.body = html;
    } else if(ctx.url==='/bomlist' && ctx.method === 'POST') {
      let result = await generateExcel();
      let fileName = result;
      let reqPath = path.join(__dirname, staticPath);
      reqPath =  path.join(reqPath, fileName);

      let content = ''
      //判断访问地址是文件夹还是文件
      content = fs.readFileSync(reqPath);
      
      let _content = content;
      // 解析请求内容的类型
      let _mime = mimes['xlsx'];
      ctx.response.type = _mime;
      ctx.response.body = _content;
      try {
        let tempPath = './tempExcel';
        let files = fs.readdirSync(tempPath);
        files.forEach(function(file, index) {
          var curPath = tempPath + '/' + file;
          fs.unlinkSync(curPath);
        });
      } catch(err) {
        console.error(err);
      }
    } 
    else{
      ctx.body='<h1>404!</h1>';
    }
});

app.listen(3000,()=>{
    console.log('server is starting at port 3000');
});