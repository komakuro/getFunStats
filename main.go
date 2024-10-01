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
	Amount    string
	Condition string
	GetMonth  string
}

// 支援者情報の構造体
type payStats struct {
	UserName  string
	PayTime   string
	PayAmount string
}

func loadConfig() settings {

	//フォーマットを開く
	f, err := excelize.OpenFile("入出力フォーマット.xlsx")
	if err != nil {
		panic("loadConfig excelize.OpenFile err:" + err.Error())
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
	cfg.CreatorId, _ = f.GetCellValue("設定", "C6")
	cfg.LoginId, _ = f.GetCellValue("設定", "C4")
	cfg.Password, _ = f.GetCellValue("設定", "C5")
	cfg.Duration, _ = f.GetCellValue("設定", "C10")
	cfg.Amount, _ = f.GetCellValue("設定", "C11")
	cfg.Condition, _ = f.GetCellValue("設定", "C12")
	cfg.GetMonth, _ = f.GetCellValue("設定", "C13")

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
	//var screenshotNum int = 1
	sets := loadConfig()

	// chromeを起動
	driver := agouti.ChromeDriver()
	driver.Start()
	defer driver.Stop() // chromeを終了

	//支援者一覧ページを開く
	page, _ := driver.NewPage()
	page.Navigate("https://" + sets.CreatorId + ".fanbox.cc/manage/relationships")
	//page.Screenshot("Screenshot" + ItoS(&screenshotNum) + ".png")

	//ログインを行う
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

	//支援者一覧を取得
	supportUsers := page.AllByClass("Row__UserWrapper-sc-1xb9lq9-1")
	supportUsersCount, _ := supportUsers.Count()
	fmt.Println("supportUsersCount", supportUsersCount)

	f := getFile("output.csv")
	defer f.Close()

	var payStatsList [][]string

	//各支援者ページに遷移して支払い日時と支払金額をスライスに格納していく
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

			if j%2 == 0 {
				oneLine.PayTime = txt
			}

			if j%2 == 1 {
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

		oneLineList := [][]string{{oneLine.UserName, oneLine.PayTime, oneLine.PayAmount}}

		//取得した支援者名、支払い日時、支払金額をスライスに格納
		payStatsList = append(payStatsList, oneLineList)

		time.Sleep(1 * time.Second)

		page.Back()

		//page.Screenshot("Screenshot" + ItoS(&screenshotNum) + ".png")
	}

	//スライスの情報を整理するためのマップを作成
	var userPaySeqMap map[string]map[string]string

	//スライスに格納された内容の数だけマップに情報を格納
	for i := 0; i < len(payStatsList); i++ {

		var tmpPayUser string = payStatsList[i][0]
		var tmpPaySeqMap map[string]string

		//マップ内に該当の支援者名が存在するか確認し、存在しなければ格納用のマップを作成
		if _, ok := userPaySeqMap[tmpPayUser]; ok {
			tmpPaySeqMap = userPaySeqMap[tmpPayUser]
		} else {
			//var tmpPaySeqMap map[string]string
		}

		var tmpPayDate string = payStatsList[i][1]
		var tmpPayAmount string = payStatsList[i][2]

		//マップ内に該当の支払い月が存在するか確認し、存在すれば支払金額を合算
		if _, ok := userPaySeqMap[tmpPayDate]; ok {
			tmpPaySeqMap[tmpPayDate] = tmpPaySeqMap[tmpPayDate] + tmpPayAmount
		} else {
			tmpPaySeqMap[tmpPayDate] = tmpPayAmount
		}

		userPaySeqMap[tmpPayUser] = tmpPaySeqMap[tmpPayUser]
	}

	var counter int = 0
	var userResultMap map[string]bool
	var checkTime time = time.Now()
	var durationTime int = strings.TrimRight(sets.Duration, "+")

	//支援者ごとの支払い情報から入力条件を満たす支援者を判定
	for i := 0; i < len(userPaySeqMap); i++ {
		for iYearMonth := checkTime; iYearMonth < checkTime-sets.GetMonth+1; iYearMonth-- {
			if sets.Condition == "継続" {
				iYearMonth = time.Date()

				if userPaySeqMap[user] == sets.Amount {
					counter = counter + 1
				} else {
					if strings.HasSuffix(sets.Duration, "+") == true {
						if counter >= durationTime {
							userResultMap[user] = true
						} else {
							userResultMap[user] = false
						}
					} else {
						if counter%durationTime == 0 {
							userResultMap[user] = true
						} else {
							userResultMap[user] = false
						}
					}
					break
				}
			} else if sets.Condition == "累積" {
				if userPaySeqMap[i] == sets.Amount {
					counter = counter + 1
				}
			}
		}

		if sets.Condition == "累積" {
			if sets.Duration == "+" {
				if counter >= durationTime {
					userResultMap[user] = true

				} else {
					userResultMap[user] = false

				}
			} else {
				if counter%durationTime == 0 {
					userResultMap[user] = true

				} else {
					userResultMap[user] = false

				}
			}
		}
	}

	//フォーマットを開く
	f, err := excelize.OpenFile("入出力フォーマット.xlsx")
	if err != nil {
		panic("loadConfig excelize.OpenFile err:" + err.Error())
	}

	var outputSheetName string = "リスト"

	//まずは出力欄の情報をクリア

	var userColoumId int = 2
	var resultColoumId int = 3
	var userRowId int = 3
	var yearMonthRowId int = 2
	var firstIter bool = true
	var yearMonthColoumId int = 4

	//判定した情報をExcelに出力していく
	for i := 0; i < len(userPaySeqMap); i++ {
		f.SetCellValue(outputSheetName, excelize.CoordinatesToCellName(userColoumId, userRowId+i), userPaySeqMap[i])

		if userResultMap[i] == true {
			f.SetCellValue(outputSheetName, excelize.CoordinatesToCellName(resultColoumId, userRowId+i), "対象")

		}

		for iYearMonth := checkTime - sets.GetMonth + 1; checkTime-sets.GetMonth < iYearMonth; iYearMonth++ {
			if firstIter == true {
				f.SetCellValue(outputSheetName, excelize.CoordinatesToCellName(yearMonthColoumId, yearMonthRowId), iYearMonth)

			}

			if _, ok := userPaySeqMap[i]; ok {
				f.SetCellValue(outputSheetName, excelize.CoordinatesToCellName(yearMonthColoumId, userRowId+i), userPaySeqMap[i][iYearMonth])

			}
			yearMonthColoumId = yearMonthColoumId + 1

		}
		userRowId = userRowId + 1
		firstIter = false

	}
}
