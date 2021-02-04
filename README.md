# 自动化截取刷卡记录

## Description
指定日期，指定员工姓名或工号，截取刷卡记录。

## 具体步骤
1. 登陆网站
[公司打卡记录查询](http://twhratsql.whq.wistron/OGWeb/LoginForm.aspx)    
用户名<kbd>S0808001</kbd>&nbsp;&nbsp;密码 <kbd>S0808001</kbd>
2. 获取登录后的cookie
3. 打开查询页面，获取查询需要的字段
4. 根据指定值查询打卡记录并将结果输出
   1. 将查询结果整合保存
   2. 判断状态
5. 继续查询下一位
6. 查询结束，保存成excel

## 判断状态
   * 正常
   * 异常
     * 請假 （迟到 + 工时不足）
     * 遲到  上班打卡时间晚于08:50:59
     * 早退  下班打卡时间早于16:50:00
     * 工時不足 上班时间和下班时间之差小于9小时
     * 只刷一次 默认认为是下班未打卡

## 生成的Excel格式
目前是每个人的记录放到一个sheet    
代码中也提供了每个人生成一个excel的方法

## How to use
```
git clone https://github.com/juanyilxc/kq.git
node index.js
```

## Dependencies
`"node-xlsx": "^0.16.1"`

## 其他
forked from [alex632/kq](https://github.com/alex632/)     

Alex的v2.0 [alex632/kq2](https://github.com/alex632/kq2)      
花春榮v2.0 [s1001001/kq](https://github.com/s1001001/kq)    
吴永辉的v2.0 [Yonghui1208/kq](https://github.com/Yonghui1208/kq)    
王洪元的v2.0 [Hongyuan-001/kq](https://github.com/Hongyuan-001/kq)    
[mamoru1314/NodeJSPractice](https://github.com/mamoru1314/NodeJSPractice)    
[dingdc/goalkeeper](https://github.com/dingdc/goalkeeper)    

# Insights
    Event driven & asynchronism are good. However they are somewhat not easy to comprehend.
    Customized Web Crawler.
