const Koa = require('koa')
const app = new Koa()
const axios = require('axios')
const Excel = require('exceljs')

// 查询数据
async function queryData() {
    let dataList = [];
    for (let i = 0; i < 12 * 4; i += 12) {
        let url = `https://www.iqianjin.com/plan/detail/buyRecord?planId=12180&sid=bc90eb04f9&pageIndex=${i}&pageSize=12&_=1543920341612`
        console.log(url)
        let data = await axios.get(url)
        dataList = dataList.concat(data.data.bean.list)
    }
    makeExcel(dataList)
}

//计算总金额
function countAmount(data) {
    let amount = 0;
    data.forEach((item, index) => {
        amount += item.amount;
    })
    return amount.toFixed(0)
}

// 生成表格
async function makeExcel(dataList) {
    let amount = await countAmount(dataList)
    //create a workbook
    let workbook = new Excel.Workbook()

    //add header
    let ws = workbook.addWorksheet("aqj")
    ws.columns = [{
            header: '用户名',
            key: 'name',
            width: 30
        },
        {
            header: `金额${amount}元`,
            key: 'amount',
            width: 30
        },
        {
            header: '创建时间',
            key: 'time',
            width: 50
        }
    ]
    ws.getColumn('amount').alignment = { horizontal: 'left' }

    dataList.forEach((item, index) => {
        ws.addRow([item.userName, item.amount, item.createTime]);
    })

    workbook.xlsx.writeFile('爱钱进.xlsx')
        .then(() => {
            console.log('生成 xlsx');
        })
        .catch(error => {
            console.log(error);
        })
}

queryData()

app.listen(3000, () => {
    console.log('listen on 3000');
})
