//用go语言抓取 豆瓣电影top250

package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"regexp"
	"strconv"
	"time"

	"github.com/tealeg/xlsx"
	"strings"
	"EmailTool"
	"bytes"
)

const (
	FileName    = "Top250.xlsx"
)



//定义新的数据类型
type Spider struct {
	url    string
	header map[string]string
}

//定义 Spider get的方法
func (keyword Spider) get_html_header() string {
	client := &http.Client{}
	req, err := http.NewRequest("GET", keyword.url, nil)
	if err != nil {
	}
	for key, value := range keyword.header {
		req.Header.Add(key, value)
	}

	resp, err := client.Do(req)
	if err != nil {
		log.Fatal(err)
	}

	defer resp.Body.Close()

	body, err := ioutil.ReadAll(resp.Body)

	if err != nil {
		log.Fatal(err)
	}

	return string(body)

}
func parse() {
	header := map[string]string{
		"Host":                      "movie.douban.com",
		"Connection":                "keep-alive",
		"Cache-Control":             "max-age=0",
		"Upgrade-Insecure-Requests": "1",
		"User-Agent":                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36",
		"Accept":                    "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
		"Referer":                   "https://movie.douban.com/top250",
	}

	//循环每页解析并把结果写入excel
	for i := 0; i < 10; i++ {
		fmt.Println("正在抓取第" + (strconv.Itoa(i + 1)) + "页......")
		url := "https://movie.douban.com/top250?start=" + strconv.Itoa(i*25) + "&filter="
		spider := &Spider{url, header}
		html := spider.get_html_header()

		//评价人数
		pattern2 := `<span>(.*?)评价</span>`
		rp2 := regexp.MustCompile(pattern2)
		find_txt2 := rp2.FindAllStringSubmatch(html, -1)

		//评分
		pattern3 := `property="v:average">(.*?)</span>`
		rp3 := regexp.MustCompile(pattern3)
		find_txt3 := rp3.FindAllStringSubmatch(html, -1)

		//电影名称
		pattern4 := `<img width="100" alt="(.*?)" src=`
		//pattern4 := `<span class="title">(.*?)</span>`
		rp4 := regexp.MustCompile(pattern4)
		find_txt4 := rp4.FindAllStringSubmatch(html, -1)

		file, err := xlsx.OpenFile(FileName)
		if err != nil {
			panic(err)
		}

		//  打印全部数据和写入excel文件
		for i := 0; i < len(find_txt2); i++ {
			fmt.Printf("%s %s %s\n", find_txt4[i][1], find_txt3[i][1], find_txt2[i][1])

			first := file.Sheets[0]
			row := first.AddRow()
			row.SetHeightCM(1)
			cell := row.AddCell()
			cell.Value = find_txt4[i][1]
			cell = row.AddCell()
			cell.Value = find_txt3[i][1]
			cell = row.AddCell()
			cell.Value = find_txt2[i][1]
		}

		err = file.Save(FileName)
		if err != nil {
			panic(err)
		}
	}
}

//创建xlsx
func CreateXlsx() {
	file := xlsx.NewFile()
	sheet, errc := file.AddSheet("Top250")
	if errc != nil {
		panic(errc)
	}
	row := sheet.AddRow()
	row.SetHeightCM(1) //设置每行的高度
	//"电影名称", "评分", "评价人数"
	cell := row.AddCell()
	cell.Value = "电影名称"
	cell = row.AddCell()
	cell.Value = "评分"
	cell = row.AddCell()
	cell.Value = "评价人数"

	err := file.Save(FileName)
	if err != nil {
		panic(err)
	}
}

func getCellValues(r *xlsx.Row) (cells []string)  {
	for _, cell := range r.Cells {
		/* 去除换行和空格 */
		txt := strings.Replace(strings.Replace(cell.Value, "\n", "", -1)," ", "", -1)
		/* 使用append函数拼接 */
		cells = append(cells, txt)
	}
	return
}

func sendEmail() {
	xlFile, err := xlsx.OpenFile(FileName)
	if err != nil {
		log.Fatalln("err:", err.Error())
	}

	var buffer bytes.Buffer
	for _, sheet := range xlFile.Sheets {
		for _, row := range sheet.Rows {
			cells := getCellValues(row)
			fmt.Println(cells)
			buffer.WriteString(fmt.Sprintf("%s\n\n", strings.Join(cells, "   ")))
		}
	}

	email := EmailTool.NewEmail("397027757@qq.com;2437319854@qq.com;kidzss@163.com;",
		"豆瓣电影Top250", buffer.String())

	err = EmailTool.SendEmail(email)
	if err!= nil {
		fmt.Println(err)
	}
}

func main() {
	// get current time
	t1 := time.Now()
	//创建xlsx
	CreateXlsx()
	//抓取数据
	parse()

	elapsed := time.Since(t1)
	fmt.Println("爬虫结束,总共耗时: ", elapsed)
	//发送邮件
	sendEmail()
}
