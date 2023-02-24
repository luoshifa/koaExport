const Koa = require('koa');
const Router = require('@koa/router');
const XLSX = require('xlsx')

const app = new Koa();
const router = new Router();
let json = [
    {

    },
    {
        "姓名": "张三",
        "性别": "男",
        "年龄": 18
    },
    {
        "姓名": "李四",
        "性别": "女",
        "年龄": 19,
        "籍贯": "江西"
    },
    {
        "姓名": "王二麻",
        "性别": "未知",
        "年龄": 20
    },
    {}
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
    let sheet = XLSX.utils.json_to_sheet(json)
    sheet['!cols'] = [
        {wch: 20},{wch: 30},{wch: 40},{wch:40}
    ]
    sheet['!rows'] = [
        {hpt:60},{hpt: 40}
    ]
    sheet['C5'] = { t:'n',f:"SUM(C2:C4)",F:"C5:C5" };
    sheet['A2'] = {t: 's', v: '新辉眼镜有限公司'}
    sheet['!merges'] = [
        {s:{r:1,c:0},e:{r:1,c:3}}
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