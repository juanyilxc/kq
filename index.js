//
// 自动化截取刷卡资料
//

'use strict';	// Whole-script strict mode applied.

const http = require('http');   // NOTE: import default module
const fs = require('fs');       // NOTE: import default module
const querystring = require('querystring'); // NOTE: import default module
const xlsx = require('node-xlsx')

//
// Step 1: Open login page to get cookie 'ASP.NET_SessionId' and hidden input '_ASPNetRecycleSession'.
//
var _ASPNET_SessionId;
var _ASPNetRecycleSession;

function openLoginPage() {

    function callback(response) {
        let chunks = [];
        response.addListener('data', (chunk) => {
            chunks.push(chunk);
        });
        response.on('end', () => {
            let buff = Buffer.concat(chunks);
            let html = buff.toString();
            if (response.statusCode === 200) {
                let fo = fs.createWriteStream('tmp/step1-LoginPage.html');
                fo.write(html);
                fo.end();
                let cookie = response.headers['set-cookie'][0];
                let patc = new RegExp('ASP.NET_SessionId=(.*?);');
                let mc = patc.exec(cookie);
                if (mc) {
                    _ASPNET_SessionId = mc[1];
                    console.log(`Cookie ASP.NET_SessionId: ${_ASPNET_SessionId}`);
                }
                let patm = new RegExp('<input type="hidden" name="_ASPNetRecycleSession" id="_ASPNetRecycleSession" value="(.*?)" />');
                let mm = patm.exec(html);
                if (mm) {
                    _ASPNetRecycleSession = mm[1];
                    console.log(`Element _ASPNetRecycleSession: ${_ASPNetRecycleSession}`);
                }
                console.log('Step1 login page got.\n');
                login();
            } else {
                let msg = `Step1 HTTP error: ${response.statusMessage}`;
                console.error(msg);
            }
        });
    }

    let req = http.request("http://twhratsql.whq.wistron/OGWeb/LoginForm.aspx", callback);

    req.on('error', e => {
        let msg = `Step1 Problem: ${e.message}`;
        console.error(msg);
    });

    req.end();
}

//
// Step 2: POST data to login to get cookie 'OGWeb'.
//
var OGWeb;

function login() {

    function callback(response) {
        let chunks = [];
        response.addListener('data', (chunk) => {
            chunks.push(chunk);
        });
        response.on('end', () => {
            let buff = Buffer.concat(chunks);
            let html = buff.toString();
            if (response.statusCode === 302) {
                let fo = fs.createWriteStream('tmp/step2-login.html');
                fo.write(html);
                fo.end();
                let cookie = response.headers['set-cookie'][0];
                let patc = new RegExp('OGWeb=(.*?);');
                let mc = patc.exec(cookie);
                if (mc) {
                    OGWeb = mc[1];
                    console.log('Cookie OGWeb got.');
                }
                console.log('Step2 done.\n');
                step3();
            } else {
                let msg = `Step2 HTTP error: ${response.statusMessage}`;
                console.error(msg);
            }
        });
    }

    let postData = querystring.stringify({
        '__ctl07_Scroll': '0,0',
        '__VIEWSTATE': '/wEPDwULLTEyMTM0NTM5MDcPFCsAAmQUKwABZBYCAgMPFgIeBXN0eWxlBTFiZWhhdmlvcjp1cmwoL09HV2ViL3RxdWFya19jbGllbnQvZm9ybS9mb3JtLmh0Yyk7FhACCA8UKwAEZGRnaGQCCg8PFgIeDEVycm9yTWVzc2FnZQUZQWNjb3VudCBjYW4gbm90IGJlIGVtcHR5LmRkAgwPDxYCHwEFGlBhc3N3b3JkIGNhbiBub3QgYmUgZW1wdHkuZGQCDQ8PFgIeB1Zpc2libGVoZGQCDg8UKwAEZGRnaGQCEg8UKwADDxYCHgRUZXh0BSlXZWxjb21lIFRvIOe3r+WJteizh+mAmuiCoeS7veaciemZkOWFrOWPuGRkZ2QCFA8UKwADDxYCHwMFK0Jlc3QgUmVzb2x1dGlvbjoxMDI0IHggNzY4OyBJRSA2LjAgb3IgYWJvdmVkZGdkAhsPFCsAAmQoKWdTeXN0ZW0uRHJhd2luZy5Qb2ludCwgU3lzdGVtLkRyYXdpbmcsIFZlcnNpb249Mi4wLjAuMCwgQ3VsdHVyZT1uZXV0cmFsLCBQdWJsaWNLZXlUb2tlbj1iMDNmNWY3ZjExZDUwYTNhBDAsIDBkGAEFHl9fQ29udHJvbHNSZXF1aXJlUG9zdEJhY2tLZXlfXxYCBQVjdGwwNwUITG9naW5CdG6vo0TFNrmm9RKH7uSQ+NY2OXccyA==',
        '__VIEWSTATEGENERATOR': 'F163E3A2',
        '_PageInstance': '1',
        '__EVENTVALIDATION': '/wEWBAK20LBAAsiTss0OArOuiN0CArmtoJkDPmmwqug37xjPhGglEwK8JU9zleg=',
        'UserPassword': 'S0808001',
        'UserAccount': 'S0808001',
        'LoginBtn.x': '74',
        'LoginBtn.y': '10',
        '_ASPNetRecycleSession': _ASPNetRecycleSession
    });
    //console.log(postData);
    let req = http.request({
        hostname: "twhratsql.whq.wistron",
        path: "/OGWeb/LoginForm.aspx",
        method: "POST",
        headers: {
            'Cookie': 'ASP.NET_SessionId=' + _ASPNET_SessionId,   // NOTED.
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': Buffer.byteLength(postData)
        }
    }, callback);

    req.on('error', e => {
        let msg = `Step2 Problem: ${e.message}`;
        console.error(msg);
    });

    req.write(postData);
    req.end();
}

//
// Step 3: Open EntryLogQueryForm.aspx page to get hidden input '_ASPNetRecycleSession', '__VIEWSTATE' and '__EVENTVALIDATION'.
//
var __VIEWSTATE = '';
var __EVENTVALIDATION = '';

function step3() {

    function callback(response) {
        let chunks = [];
        response.addListener('data', (chunk) => {
            chunks.push(chunk);
        });
        response.on('end', () => {
            let buff = Buffer.concat(chunks);
            let html = buff.toString();
            if (response.statusCode === 200) {
                let fo = fs.createWriteStream('tmp/step3.html');
                fo.write(html);
                fo.end();
                let patm = new RegExp('<input type="hidden" name="_ASPNetRecycleSession" id="_ASPNetRecycleSession" value="(.*?)" />');
                let mm = patm.exec(html);
                if (mm) {
                    _ASPNetRecycleSession = mm[1];
                    console.log(`Element _ASPNetRecycleSession: ${_ASPNetRecycleSession}`);
                }
                let patv = new RegExp('<input type="hidden" name="__VIEWSTATE" id="__VIEWSTATE" value="(.*?)"');
                let mv = patv.exec(html);
                if (mv) {
                    __VIEWSTATE = mv[1];
                    console.log('Element __VIEWSTATE got');
                }
                let pate = new RegExp('<input type="hidden" name="__EVENTVALIDATION" id="__EVENTVALIDATION" value="(.*?)"');
                let me = pate.exec(html);
                if (me) {
                    __EVENTVALIDATION = me[1];
                    console.log('Element __EVENTVALIDATION got');
                }
                console.log('Step3 done.\n');
                askAll();
            } else {
                let msg = `Step3 HTTP error: ${response.statusMessage}`;
                console.error(msg);
            }
        });
    }

    let req = http.request({
        hostname: "twhratsql.whq.wistron",
        path: "/OGWeb/OGWebReport/EntryLogQueryForm.aspx",
        //method: "GET",    // Default can be omitted.
        headers: {
            'Cookie': `ASP.NET_SessionId=${_ASPNET_SessionId}; OGWeb=${OGWeb}`  // important
        }
    }, callback);

    req.on('error', e => {
        let msg = `Step3 Problem: ${e.message}`;
        console.error(msg);
    });

    req.end();
}

//记录查询到了那天用于判断是否是工作日
var begDate;
//用来保存每个人的打开记录
var results = [];
//用来保存所有人的打卡记录
var temp = [];

//
//将Step 4封装到Promise中
//
/**
 * 参数定义和Step 4一样
 */
const pr = (beginDate, endDate, employeeIdOrName, nextPage) => {
    begDate = endDate;
    results = [];
    return new Promise((resolve, reject) => {
        inquire(beginDate, endDate, employeeIdOrName, nextPage, (err, data) => {
            if (err) {
                return reject(err);
            }
            resolve(data);
        });
    });
}

//
// Step 4: POST data to inquire.
//
/**
 * 截取某人的刷卡资料。
 * @param {*} beginDate 开始日期
 * @param {*} endDate 截止日期
 * @param {*} employeeIdOrName 工号或名字
 * @param {*} nextPage if go to next page
 * @param {*} nextStep 完成后调用此function
 */
function inquire(beginDate, endDate, employeeIdOrName, nextPage, nextStep) {

    function callback(response) {
        let chunks = [];
        response.addListener('data', (chunk) => {
            chunks.push(chunk);
        });
        response.on('end', () => {
            let buff = Buffer.concat(chunks);
            let html = buff.toString();
            if (response.statusCode === 200) {
                let result = parseKQ(html);
                let fo = fs.createWriteStream(`tmp/step4-inquire-${employeeIdOrName}-${result.curPage}.html`);
                fo.write(html);
                fo.end();
                if (result.curPage < result.numPages) {
                    inquire(beginDate, endDate, employeeIdOrName, true, nextStep);
                } else {
                    console.log(`Inquiry about ${employeeIdOrName} is done.`);
                    setState(results);
                    temp.push(results);//将结果保存到temp
                    //toSignelExcel(results);
                    if (nextStep) {   // If provided.
                        nextStep();
                    }
                }
            } else {
                console.error(`Inquiry HTTP error: ${response.statusMessage}`);
            }
        });
    }

    var beginTime = '0:00';
    var endTime = '23:59';

    let postObj = {
        'TQuarkScriptManager1': 'QueryResultUpdatePanel|QueryBtn',
        'TQuarkScriptManager1_HiddenField': ';;AjaxControlToolkit, Version=1.0.20229.20821, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:en-US:c5c982cc-4942-4683-9b48-c2c58277700f:411fea1c:865923e8;;AjaxControlToolkit, Version=1.0.20229.20821, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:en-US:c5c982cc-4942-4683-9b48-c2c58277700f:91bd373d:d7d5263e:f8df1b50;;AjaxControlToolkit, Version=1.0.20229.20821, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:en-US:c5c982cc-4942-4683-9b48-c2c58277700f:e7c87f07:bbfda34c:30a78ec5;;AjaxControlToolkit, Version=1.0.20229.20821, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:en-US:c5c982cc-4942-4683-9b48-c2c58277700f:9b7907bc:9349f837:d4245214;;AjaxControlToolkit, Version=1.0.20229.20821, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:en-US:c5c982cc-4942-4683-9b48-c2c58277700f:e3d6b3ac;',
        '__ctl07_Scroll': '0,0',
        '__VIEWSTATEGENERATOR': 'A21EDEFC',
        '_ASPNetRecycleSession': _ASPNetRecycleSession,
        '__VIEWSTATE': __VIEWSTATE,
        '_PageInstance': 26,
        '__EVENTVALIDATION': __EVENTVALIDATION,
        'AttNoNameCtrl1$InputTB': '上海欽江路',
        'BeginDateTB$Editor': beginDate,
        'BeginDateTB$_TimeEdit': beginTime,
        'EndDateTB$Editor': endDate,
        'EndDateTB$_TimeEdit': endTime,
        'EmpNoNameCtrl1$InputTB': employeeIdOrName
    };
    if (nextPage) {
        postObj['GridPageNavigator1$NextBtn'] = 'Next Page';
    } else {
        postObj['QueryBtn'] = 'Inquire';
    }

    let postData = querystring.stringify(postObj);

    let req = http.request({
        hostname: "twhratsql.whq.wistron",
        path: "/OGWeb/OGWebReport/EntryLogQueryForm.aspx",
        method: "POST",
        headers: {
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; MAARJS)',	// mimic IE 11 // important
            'X-MicrosoftAjax': 'Delta=true',    // important
            'Cookie': `ASP.NET_SessionId=${_ASPNET_SessionId}; OGWeb=${OGWeb}`,  // important
            'Content-Type': 'application/x-www-form-urlencoded',
            'Content-Length': Buffer.byteLength(postData)
        }
    }, callback);

    req.on('error', e => {
        console.error(`Step4 Problem: ${e.message}`);
    });

    req.end(postData);
}


/**
 * Parse the input html to get 刷卡 data.
 * @param {*} html 
 * @return number of current page and number of total pages.
 */
function parseKQ(html) {
    // Get number of pages.
    let curPage = 1;
    let numPages = 1;
    let rexTotal = new RegExp('<span id="GridPageNavigator1_CurrentPageLB">(.*?)</span>[^]*?<span id="GridPageNavigator1_TotalPageLB">(.*?)</span>');
    let match = rexTotal.exec(html);
    if (match) {
        curPage = parseInt(match[1]);
        numPages = parseInt(match[2]);
        console.log(`Page: ${curPage} / ${numPages}`);
    }

    // Update __VIEWSTATE __EVENTVALIDATION
    let rexVS = new RegExp("__VIEWSTATE[\|](.*?)[\|]");
    let matVS = rexVS.exec(html);
    if (matVS) {
        __VIEWSTATE = matVS[1];
    }
    let rexEV = new RegExp("__EVENTVALIDATION[\|](.*?)[\|]");
    let matEV = rexEV.exec(html);
    if (matEV) {
        __EVENTVALIDATION = matEV[1];
    }
    //倒计时 记录工作日
    let workday = new Date(begDate);
    // Print 刷卡 data
    console.log(`/Department  /EID  /Name  /Clock Time`);
    while (true) {
        let rex = new RegExp('<td>(.*?)</td><td>&nbsp;</td><td><.*?>(.*?)</a></td><td>(.*?)</td><td>.*?</td><td>(.*?)</td>',
            'g');   // NOTE: 'g' is important
        let m = rex.exec(html);
        if (m) {
            console.log(`${m[1]} ${m[2]} ${m[3]} ${m[4]}`);//部门 工号 姓名 打卡日期时间
            let datetime = new Date(m[4]);
            //m的日期和计数日期比对，当计数大于m的日期，就判断计数日期是星期几，如果是周一~周五，表示当天未刷卡，加入一笔记录（后续会判断为请假，实际还可能是忘刷卡，放假，此处不考虑此种情况）
            //result数据格式
            //部门 工号 姓名 日期 第一次打卡时间 最后一次打卡时间 状态（请假 迟到等）
            //while(workday.getMonth() != datetime.getMonth() && workday.getDate() != datetime.getDate()){
            while (workday.toLocaleDateString("chinese",{month: '2-digit',day: '2-digit',year: 'numeric'}) > datetime.toLocaleDateString("chinese",{month: '2-digit',day: '2-digit',year: 'numeric'})) {
                if (workday.getDay() > 0 && workday.getDay() < 6) {
                    //results.push([m[1],m[2],m[3],workday,'','','']);
                    results.push([m[1], m[2], m[3], new Date(workday.toLocaleDateString()), '', '', '']);
                    //break;
                }
                workday.setDate(workday.getDate() - 1);
            }

            //let rex = new RegExp('([0-9]{1,2})\/([0-9]{1,2})\/[0-9]{4}\\s([0-9]{1,2}\:[0-9]{1,2}\:[0-9]{1,2})'); 
            //let res = rex.exec(m[4]);//月份 日期 时间
            //当result有数据时，判断此笔m是否和前一笔已经存在的数据是同一天
            //如果为同一天，则放到result中的第一次打卡时间字段（因为数据是倒叙，所以第二次出现的为上班打卡时间）
            if (results.length > 0) {
                let lastdate = results[results.length - 1][3];

                if (lastdate.getMonth() == datetime.getMonth() && lastdate.getDate() == datetime.getDate()) {
                    results[results.length - 1][4] = datetime;
                    workday.setDate(workday.getDate() + 1);//同一天的打卡记录，为了防止if判断外面把workday多减1天
                }
                else
                    results.push([m[1], m[2], m[3], new Date(datetime.toLocaleDateString()), '', datetime, '']);
            } else {
                results.push([m[1], m[2], m[3], new Date(datetime.toLocaleDateString()), '', datetime, '']);
            }
            workday.setDate(workday.getDate() - 1);
            //toExcel(m);
            html = html.substr(rex.lastIndex);

        } else {
            //
            //workday.setDate(workday.getDate() - 1);
            begDate = workday;
            break;
        }
    }
    return { curPage: curPage, numPages: numPages };
}

//判断状态
function setState(res) {
    res.forEach(e => {
        if (e[4] == '') {
            e[6] += '请假';
        } else {
            if (e[5] == '')
                e[6] += '下班未打卡';
            //console.log(e[4]);
            if (e[4] != '' && e[4].toLocaleTimeString('chinese', { hour12: false,hour:"2-digit" ,minute:"2-digit",second:"2-digit"}) > "08:50:59")
                e[6] += '迟到';
            if (e[5] != '' && e[5].toLocaleTimeString('chinese', { hour12: false,hour:"2-digit" ,minute:"2-digit",second:"2-digit"}) < "16:50:00")
                e[6] += '早退';
            if (e[4] != '' && e[5] != '' && (e[5].getTime() - e[4].getTime()) / (1000 * 60 * 60) < 9)
                e[6] += '工时不足';
            if (e[6] == '')
                e[6] = '正常';
        }
    });
    //迟到+工时不足=>请假
    res.forEach(e => {
        if (e[6].indexOf("工时不足") >= 0 && e[6].indexOf("迟到") >= 0)
            e[6] = '请假';
    })
}

//利用node xlsx导成excel
function toExcel(mt) {
    let datas = [];
    let title = ['部门', '工号', '姓名', '日期', '上班时间', '下班时间', '状态']; //这是第一行 俗称列名 
    //rows是个从数据库里面读出来的数组，大家就把他当成一个普通的数组就ok
    mt.forEach( m => {
        let data = []; // 其实最后就是把这个数组写入excel 
        data.push(title); // 添加完列名 下面就是添加真正的内容了
        m.forEach((e) => {
            let arrInner = [];
            arrInner.push(e[0]);
            arrInner.push(e[1]);
            arrInner.push(e[2]);
            arrInner.push(e[3].toLocaleDateString());
            arrInner.push(e[4] == '' ? '' : e[4].toLocaleTimeString('chinese', { hour12: false }));
            arrInner.push(e[5] == '' ? '' : e[5].toLocaleTimeString('chinese', { hour12: false }));
            arrInner.push(e[6]);
            data.push(arrInner); //data中添加的要是数组，可以将对象的值分解添加进数组，例如：['1','name','上海']
        });

        let t = {
            name: data[1][1],
            data: data
        }

        datas.push(t);
    });
    writeXls(datas);
};

function writeXls(datas) {
    let buffer = xlsx.build(datas);
    fs.writeFileSync('the_content.xlsx', buffer, { 'flag': 'w' });//生成excel the_content是excel的名字，大家可以随意命名
    
    // fs.access('./the_content.xlsx',fs.constants.F_OK,err=>{
    //     if(err)
    //         fs.writeFileSync('./the_content.xlsx', buffer, { 'flag': 'w' });//生成excel the_content是excel的名字，大家可以随意命名
    //     else
    //         fs.appendFile('./the_content.xlsx', buffer, { 'flag': 'w' },(err) => {
    //             if (err) throw err;
    //             console.log('数据已被追加到文件');
    //         })
    // })
           
}

//每个
function toSignleExcel(m) {
    let title = ['部门', '工号', '姓名', '日期', '上班时间', '下班时间', '状态']; //这是第一行 俗称列名 
    //rows是个从数据库里面读出来的数组，大家就把他当成一个普通的数组就ok
    let data = []; // 其实最后就是把这个数组写入excel 
    data.push(title); // 添加完列名 下面就是添加真正的内容了
    m.forEach((e) => {
        let arrInner = [];
        arrInner.push(e[0]);
        arrInner.push(e[1]);
        arrInner.push(e[2]);
        arrInner.push(e[3].toLocaleDateString());
        arrInner.push(e[4] == '' ? '' : e[4].toLocaleTimeString('chinese', { hour12: false }));
        arrInner.push(e[5] == '' ? '' : e[5].toLocaleTimeString('chinese', { hour12: false }));
        arrInner.push(e[6]);
        data.push(arrInner); //data中添加的要是数组，可以将对象的值分解添加进数组，例如：['1','name','上海']
    }); 
    writesignleXls(data);
};

function writesignleXls(datas) {
    let buffer = xlsx.build([
        {
            name: datas[1][1],
            data: datas
        }
    ]);
    fs.writeFileSync(`./result-${datas[1][1]}.xlsx`, buffer, { 'flag': 'w' });//生成excel the_content是excel的名字，大家可以随意命名
           
}

// 回调版
// function askAll() {
//     inquire('2020-12-24', '2021-1-11', 'john', false,
//     ()=> inquire('2020-12-28', '2021-1-11', 'S2008001', false,
//     ()=> inquire('2020-12-24', '2021-1-6', 'ANNE', false,
//     ()=> inquire('2020-12-25', '2021-1-08', 'LEO MY CHEN', false,
//     ()=> inquire('2021-01-7', '2021-1-11', 'S0203002', false,
//     function() { console.log("All done.") } )))));
// }

// let promise = new Promise((resolve, reject) => {
//     // 初始化promise状态为padding
//     // 执行异步操作
//     if(异步操作成功){
//         resolve(value) // 修改promise的状态为fullfiled
//     } else {
//         reject(errMsg) // 修改promise的状态为rejected
//     }
// })

//Promise版本
async function askAll() {
    await pr('2020-12-24', '2021-1-14', 'john', false)
    await pr('2020-12-28', '2021-1-11', 'S2008001', false)
    // await pr('2020-12-24', '2021-1-6', 'ANNE', false)
    // await pr('2020-12-25', '2021-1-08', 'LEO MY CHEN', false)
    // await pr('2021-01-7', '2021-1-11', 'S0203002', false)
    toExcel(temp)
    console.log("All done.")
}

openLoginPage();    // Where it all begins.
