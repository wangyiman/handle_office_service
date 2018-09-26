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
  postData = {"userInfo":{"designId":"2018926","designUser":"布谷","customerAddress":"浦东东方城市花园"},"bomList":[{"customComputed":[{"order":"1","name":"抽屉灶具柜","brand":"志邦","sku":"","standlizeSize":{"x":"900.00","y":"556.00","z":"700.00"},"material":{"body":"白枫木-浅色","door":"金秋色"},"unit":"延米","quantity":"1.80","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"2","name":"直角转角柜","brand":"志邦","sku":"","standlizeSize":{"x":"1200.00","y":"556.00","z":"700.00"},"material":{"body":"白枫木-浅色","door":"金秋色"},"unit":"延米","quantity":"1.84","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"3","name":"双开门高柜","brand":"志邦","sku":"","standlizeSize":{"x":"900.00","y":"556.00","z":"1400.00"},"material":{"body":"白枫木-浅色","door":"金秋色"},"unit":"延米","quantity":"1.80","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"4","name":"单层上翻吊柜","brand":"志邦","sku":"","standlizeSize":{"x":"900.00","y":"296.00","z":"700.00"},"material":{"body":"白枫木-浅色","door":"金秋色"},"unit":"延米","quantity":"1.80","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"5","name":"五角转角柜","brand":"志邦","sku":"","standlizeSize":{"x":"1200.00","y":"800.00","z":"700.00"},"material":{"body":"白枫木-浅色","door":"金秋色"},"unit":"延米","quantity":"1.60","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""}]},{"plateComputed":[{"order":"","name":"收口条L形板","brand":"通配","sku":"","standlizeSize":{"x":"50.00","y":"500.00","z":"700.00"},"material":"金秋色","unit":"根数","quantity":"1.00","number":"2","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""}]},{"eleComputed":[{"order":"1","name":"JZ(T/Y)-B7501","brand":"志邦","sku":"","standlizeSize":{"x":"750.00","y":"430.00","z":"69.00"},"material":"","unit":"cm","quantity":"1.00","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"","name":"欧琳水盆OLCT405","brand":"欧琳橱柜","sku":"","standlizeSize":{"x":"830.00","y":"460.00","z":"500.45"},"material":"","unit":"cm","quantity":"1.00","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"","name":"SG854801 不锈钢多功能单槽 带龙头","brand":"志邦","sku":"","standlizeSize":{"x":"850.02","y":"480.00","z":"604.64"},"material":"","unit":"cm","quantity":"1.00","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"","name":"SG854801 不锈钢多功能单槽","brand":"志邦","sku":"","standlizeSize":{"x":"850.02","y":"480.00","z":"281.14"},"material":"","unit":"cm","quantity":"1.00","number":"2","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""}]},{"decoComputed":[{"order":"1","name":"台面","brand":"志邦","sku":"","standlizeSize":{"x":"3535.53","y":"15.00","z":"37.00"},"material":"GDB301冰晶白","unit":"米","quantity":"3.54","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"2","name":"顶线","brand":"通配","sku":"","standlizeSize":{"x":"2828.43","y":"34.37","z":"40.00"},"material":"金秋色","unit":"米","quantity":"2.83","number":"4","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""},{"order":"3","name":"脚线","brand":"通配","sku":"","standlizeSize":{"x":"2828.43","y":"15.00","z":"100.00"},"material":"铝塑拉丝","unit":"米","quantity":"2.83","number":"1","unitPrice":"","customPrice":"","discount":"","discountPrice":"","instructions":""}]}]};
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