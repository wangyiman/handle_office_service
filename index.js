const Koa = require('koa');
const fs = require('fs');
const mimes = require('./mimes')
const path = require('path')
const generateExcel = require('./utils/generateExcel');
const Excel = require('exceljs');
const koaBody = require('koa-body');
const Router = require('koa-router');

const router = new Router();
const app = new Koa();
const staticPath = './tempExcel';    

app
  .use(koaBody({
    "formLimit": "5mb",
    "jsonLimit": "50mb",
    "textLimit": "5mb"
  }))
  .use(router.routes())
  .use(router.allowedMethods());

router.get('/', async (ctx, next) => {
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
  };
  var finalStr = JSON.stringify(finnalData);
    let html =`<h1>Koa2 request post demo</h1>
        <form method="POST"  action="/bomlist">
            <input name='webSite' value=${finalStr} type='text'/>
            <br/>
            <button type="submit">submit</button>
        </form>`;
    ctx.response.type = 'text/html';        
    ctx.response.body = html;
});

router.post('/bomlist', async (ctx, next) => {
  let postData = ctx.request.body;
  console.log(postData);
  postData ={"userInfo":{"designId":"2018926","designUser":"布谷","customerAddress":"朝阳wwww"},"bomList":[{"customComputed":[]},{"decoComputed":[]},{"eleComputed":[]},{"plateComputed":[]}]};
  let result = await generateExcel(postData);
  let fileName = result;
  let reqPath = path.join(__dirname, staticPath);
  reqPath =  path.join(reqPath, fileName);

  let content = ''
  content = fs.readFileSync(reqPath);
  
  let _content = content;
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
})

app.listen(3000,()=>{
    console.log('server is starting at port 3000');
});