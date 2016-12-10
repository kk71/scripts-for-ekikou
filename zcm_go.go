// wriiten by kk.

package main

import (
    "os"
    "fmt"
    "time"
	"net/http"
	"net/url"
    "strings"
	"net/http/cookiejar"
    "github.com/tealeg/xlsx"
)

// 时间格式化的格式
const TIME_FORM = "2006-01-02"

// urls
const URL_HOST = "zcm.zcmlc.com"
const URL_LOGIN = "http://zcm.zcmlc.com/zcm/admin/login"
const URL_QUERY_ACCOUNT_PURCHASE_INFO_WITH_TIME_RANGE = "http://zcm.zcmlc.com/zcm/admin/userdetailbuy"
const URL_QUERY_ORDER_DETAIL = "http://zcm.zcmlc.com/zcm/admin/userdetailtradedetal"

// 登录信息
const USERNAME = "chenyk"
const PASSWORD = "000123"

func main() {
    xlsFileName := os.Args[1] // xls 文件名
    tels := getTels(xlsFileName)
    fmt.Println(tels)
    var cli *zcmClient = &zcmClient{}
    // cli.Login(USERNAME, PASSWORD)
    var flagToDelay int
    fmt.Print("输入延迟指数(默认为20)：")
    fmt.Scanf("%d", &flagToDelay)
    if flagToDelay==0 {
        flagToDelay = 20
        fmt.Println("使用缺省值：", flagToDelay)
    }
    var (
        inputTimeStart string
        inputTimeEnd string
    )
    fmt.Printf("开始时间(形如%s)：", TIME_FORM)
    fmt.Scanf("%s", &inputTimeStart)
    fmt.Printf("终止时间(形如%s)：", TIME_FORM)
    fmt.Scanf("%s", &inputTimeEnd)
    timeStart, _ := time.Parse(TIME_FORM, inputTimeStart)
    timeEnd, _ := time.Parse(TIME_FORM, inputTimeEnd)
    cli.GetOrderList(nil, timeStart, timeEnd)
}

// 延迟
func performDelay(flag int) {
    return
}

// zcm 客户端结构体
type zcmClient struct {
    requestClient   *http.Client
    jar             *cookiejar.Jar
}

// 读取 xls，然后处理每条记录
func readXls(fileName string) {
    file, err := xlsx.OpenFile(fileName)
    fmt.Println(file.Sheets[0].Rows[0].Cells[0].Value)
    if err!=nil {
        panic("ecel 文件读取失败。")
    }
}

// 添加 headers
func (c *zcmClient) setHeaders(request *http.Request) {
    request.Header.Set("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8")
    request.Header.Set("Accept-Encoding", "gzip, deflate, sdch")
    request.Header.Set("Accept-Language", "en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4")
    request.Header.Set("Connection", "keep-alive")
    request.Header.Set("DNT", "1")
    request.Header.Set("Host", "zcm.zcmlc.com")
    request.Header.Set("Referer", "http://zcm.zcmlc.com/zcm/admin/login")
    request.Header.Set("Upgrade-Insecure-Requests", "1")
    request.Header.Set("User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36")
}

// 登录
func (c *zcmClient) Login(username string, password string) {
    fmt.Println("尝试登陆...")
    var verifyCode string
    fmt.Printf("输入你当前的验证码：")
    n := 0
    for n<1 {n, _ = fmt.Scanf("%s", &verifyCode)}
    // construct the data-form
    formData := url.Values{}
    formData.Set("username", username)
    formData.Set("password", password)
    // initiate the cookie jar
    jar, _ := cookiejar.New(nil)
    c.jar = jar
    // initiate the http client
    c.requestClient = &http.Client{Jar: c.jar}
    // construct the request object
    req, _ := http.NewRequest("POST", URL_LOGIN, strings.NewReader(formData.Encode()))
    c.setHeaders(req)
    req.Header.Set("Content-Type", "application/x-www-form-urlencoded") // 这条 header 必须手动加，不然默认是编码成form-data形式
    // launch it!
    resp, _ := c.requestClient.Do(req)
    fmt.Printf("returned with status %d\n", resp.StatusCode)
    fmt.Println(jar)
}

// 根据 tel 查询订单列表
func (c *zcmClient) GetOrderList(row *xlsx.Row, timeStart time.Time, timeEnd time.Time) {
}

// 查询某一笔交易的详情
func (c *zcmClient) GetOrderDetail(orderId string, tel string) []string {
    var ret []string
    return ret
}