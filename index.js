const Koa = require('koa');
const Router = require('@koa/router');
const XLSX = require('xlsx')

const app = new Koa();
const router = new Router();
let arr = [
    ['新辉眼镜'],
    ['姓名','性别','年龄','籍贯'],
    ['张三','男',18,'江西'],
    ['张三','男',17,'江西'],
    ['张三','男',16,'江西'],
    ['张三','男',15,'江西'],
]
let json2 = [
    {
        "姓名": "张三2",
        "性别": "男",
        "年龄": 18
    },
    {
        "姓名": "李四2",
        "性别": "女",
        "年龄": 19,
        "籍贯": "江西"
    },
    {
        "姓名": "王二麻2",
        "性别": "未知",
        "年龄": 20
    }
]
router.get('/export', async (ctx) => {
    // 实例化一个工作簿
    let book = XLSX.utils.book_new()
    // 实例化一个Sheet
    let sheet = XLSX.utils.aoa_to_sheet(arr);
    sheet['!cols'] = [
        {wch: 20},{wch: 30},{wch: 40},{wch:40}
    ]
    sheet['!rows'] = [
        {hpt:60},{hpt: 40}
    ]
    sheet['!merges'] = [
        {s:{r:0,c:0},e:{r:0,c:3}}
    ]
    let sheet2 = XLSX.utils.json_to_sheet(json2)
    // 将Sheet写入工作簿
    XLSX.utils.book_append_sheet(book, sheet, '一班')
    XLSX.utils.book_append_sheet(book, sheet2, '二班')
    let res = XLSX.write(book,{type:'buffer'})
    ctx.set("Content-Type", "application/octet-stream")
    ctx.set("Content-Disposition", "attachment; filename=" + encodeURIComponent("文件名.xlsx"));
    ctx.body = res;
})
app.use(router.routes());
app.listen(3000)