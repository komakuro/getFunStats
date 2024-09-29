package main

import (
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/sclevine/agouti"

	"github.com/xuri/excelize/v2"
)

// 設定情報の構造体
type settings struct {
	CreatorId string
	LoginId   string
	Password  string
	Duration  string
	Amount    int
	Condition string
}

// 支援者情報の構造体
type payStats struct {
	UserName  string
	PayTime   string
	PayAmount int
}

func loadConfig() settings {

	//フォーマットを開く
	f, err := excelize.OpenFile("入出力フォーマット.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	defer f.Close()

	//f, err := os.Open("settings.json")
	//if err != nil {
	//	panic("loadConfig os.Open err:" + err.Error())
	//}
	//defer f.Close()

	var cfg settings
	//_ = json.NewDecoder(f).Decode(&cfg)

	//フォーマット内の指定されたセルの値を取得
	cfg.CreatorId = f.GetCellValue("設定", "C6")
	cfg.LoginId = f.GetCellValue("設定", "C4")
	cfg.Password = f.GetCellValue("設定", "C5")
	cfg.Duration = f.GetCellValue("設定", "C10")
	cfg.Amount = f.GetCellValue("設定", "C11")
	cfg.Condition = f.GetCellValue("設定", "C12")

	return cfg
}

func ItoS(screenshotNum *int) string {
	ret := strconv.Itoa(*screenshotNum)
	*screenshotNum += 1
	return ret
}

func getFile(filename string) *os.File {
	f, err := os.Create(filename)
	if err != nil {
		log.Fatal(err)
	}

	return f
}

func writeFile(f *os.File, writeString string) {
	d := []byte(writeString + "\n")

	// 3. 書き込み
	_, err := f.Write(d)
	if err != nil {
		log.Fatal(err)
	}
}

func main() {
	var screenshotNum int = 1
	sets := loadConfig()

	// chromeを起動
	driver := agouti.ChromeDriver()
	driver.Start()
	defer driver.Stop() // chromeを終了

	page, _ := driver.NewPage()
	page.Navigate("https://" + sets.CreatorId + ".fanbox.cc/manage/relationships")
	//page.Screenshot("Screenshot" + ItoS(&screenshotNum) + ".png")

	loginIdForm := page.AllByXPath("//*[@id=\"app-mount-point\"]/div/div/div[4]/div[1]/div[2]/div/div/div/form/fieldset[1]/label/input")
	count, _ := loginIdForm.Count()
	fmt.Println("count", count)

	loginIdForm.Fill(sets.LoginId)

	passwordForm := page.AllByXPath("//*[@id=\"app-mount-point\"]/div/div/div[4]/div[1]/div[2]/div/div/div/form/fieldset[2]/label/input")
	pasCount, _ := passwordForm.Count()
	fmt.Println("pasCount", pasCount)

	passwordForm.Fill(sets.Password)

	//page.Screenshot("Screenshot" + ItoS(&screenshotNum) + ".png")

	loginSubmit := page.AllByXPath("//*[@id=\"app-mount-point\"]/div/div/div[4]/div[1]/div[2]/div/div/div/form/button[1]")
	loginSubmitCount, _ := loginSubmit.Count()
	fmt.Println("loginSubmitCount", loginSubmitCount)

	loginSubmit.Submit()

	time.Sleep(3 * time.Second)

	//page.Screenshot("Screenshot" + ItoS(&screenshotNum) + ".png")

	time.Sleep(2 * time.Second)

	supportUsers := page.AllByClass("Row__UserWrapper-sc-1xb9lq9-1")
	supportUsersCount, _ := supportUsers.Count()
	fmt.Println("supportUsersCount", supportUsersCount)

	f := getFile("output.csv")
	defer f.Close()

	for i := 0; i < supportUsersCount; i++ {
		supportUsers.At(i).Click()

		time.Sleep(1 * time.Second)

		records := page.AllByClass("SupportTransactionSection__Td-sc-17tc9du-3")
		recordsCount, _ := records.Count()
		title, _ := page.Title()
		fmt.Println("user", title, "recordsCount", recordsCount)

		userName := strings.Split(title, "｜")[0]

		var oneLine payStats

		oneLine.UserName = userName

		for j := 0; j < recordsCount; j++ {
			record := records.At(j)
			txt, _ := record.Text()

			fmt.Println("txt", txt)

			if j == 1 {
				oneLine.PayTime = txt
			}

			if j == 2 {
				oneLine.PayAmount = txt
			}

			//oneLine += txt

			//if j%2 != 0 {
			//	writeFile(f, userName+","+oneLine)
			//	oneLine = ""
			//} else {
			//	oneLine += ","
			//}

		}

		time.Sleep(1 * time.Second)

		page.Back()

		//page.Screenshot("Screenshot" + ItoS(&screenshotNum) + ".png")
	}

}
