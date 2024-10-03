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

func readCell(f *excelize.File, sheetName string, cellPosition string) string {
	ret, err := f.GetCellValue(sheetName, cellPosition)
	if err != nil {
		panic("readCell err:" + err.Error())
	}

	return ret
}

func loadConfig() settings {

	//フォーマットを開く
	f, err := excelize.OpenFile("入出力フォーマット.xlsx")
	if err != nil {
		panic("loadConfig excelize.OpenFile err:" + err.Error())
	}
	defer f.Close()

	var cfg settings

	//フォーマット内の指定されたセルの値を取得
	cfg.CreatorId = readCell(f, "設定", "C6")
	cfg.LoginId = readCell(f, "設定", "C4")
	cfg.Password = readCell(f, "設定", "C5")
	cfg.Duration = readCell(f, "設定", "C10")
	cfg.Amount = readCell(f, "設定", "C11")
	cfg.Condition = readCell(f, "設定", "C12")
	cfg.GetMonth = readCell(f, "設定", "C13")

	return cfg
}

func ItoS(screenshotNum *int) string {
	ret := strconv.Itoa(*screenshotNum)
	*screenshotNum += 1
	return ret
}

func GetFile(filename string) *os.File {
	f, err := os.Create(filename)
	if err != nil {
		log.Fatal(err)
	}

	return f
}

func WriteFile(f *os.File, writeString string) {
	d := []byte(writeString + "\n")

	// 3. 書き込み
	_, err := f.Write(d)
	if err != nil {
		log.Fatal(err)
	}
}

func main() {
	//設定の取得
	sets := loadConfig()

	// chromeを起動
	driver := agouti.ChromeDriver()
	driver.Start()
	defer driver.Stop() // chromeを終了

	//支援者一覧ページを開く
	page, _ := driver.NewPage()
	page.Navigate("https://" + sets.CreatorId + ".fanbox.cc/manage/relationships")

	//ログインを行う
	fillForm := page.AllByClass("sc-bn9ph6-6")
	fillCount, _ := fillForm.Count()
	fmt.Println("fillCount", fillCount)

	for i := 0; i < fillCount; i++ {
		if i == 0 {
			fillForm.At(i).Fill(sets.LoginId)
		}
		if i == 1 {
			fillForm.At(i).Fill(sets.Password)
		}
	}

	loginSubmit := page.AllByClass("sc-2o1uwj-9")
	loginSubmitCount, _ := loginSubmit.Count()
	fmt.Println("loginSubmitCount", loginSubmitCount)

	loginSubmit.Submit()

	time.Sleep(1 * time.Second)

	//支援者一覧を取得
	supportUsers := page.AllByClass("Row__UserWrapper-sc-1xb9lq9-1")
	supportUsersCount, _ := supportUsers.Count()
	fmt.Println("supportUsersCount", supportUsersCount)

	var payStatsList []payStats

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
				//一行分の情報を取り終わったので、取得した支援者名、支払い日時、支払金額をスライスに格納
				payStatsList = append(payStatsList, oneLine)
				//次の一行の情報の取得処理を開始するにあたって、oneLineを初期化する
				oneLine = payStats{}
				oneLine.UserName = userName
			}

		}

		time.Sleep(1 * time.Second)

		page.Back()

	}

	//スライスの情報を整理するためのマップを作成
	var userPaySeqMap map[string]map[string]string = make(map[string]map[string]string)

	//スライスに格納された内容の数だけマップに情報を格納
	for i := 0; i < len(payStatsList); i++ {

		var tmpPayUser string = payStatsList[i].UserName
		var tmpPaySeqMap map[string]string = make(map[string]string)

		//マップ内に該当の支援者名が存在するか確認し、存在しなければ格納用のマップを作成
		if _, ok := userPaySeqMap[tmpPayUser]; ok {
			tmpPaySeqMap = userPaySeqMap[tmpPayUser]
		}

		var tmpPayDate string = payStatsList[i].PayTime
		tmpPayDate = tmpPayDate[:7]

		var tmpPayAmount string = payStatsList[i].PayAmount
		var tmpPayAmountInt, _ = strconv.Atoi(strings.TrimLeft(tmpPayAmount, "\\"))
		var tmpPayAmount2, _ = strconv.Atoi(tmpPaySeqMap[tmpPayDate])

		//マップ内に該当の支払い月が存在するか確認し、存在すれば支払金額を合算
		if _, ok := userPaySeqMap[tmpPayDate]; ok {
			var AmountSum int = tmpPayAmount2 + tmpPayAmountInt
			var AmountSumStr string = strconv.Itoa(AmountSum)
			tmpPaySeqMap[tmpPayDate] = "\\" + AmountSumStr

		} else {
			tmpPaySeqMap[tmpPayDate] = tmpPayAmount

		}

		userPaySeqMap[tmpPayUser] = tmpPaySeqMap
	}

	fmt.Println(userPaySeqMap)

	var counter int = 0
	var userResultMap = make(map[string]bool)
	var checkTime = time.Now()
	durationTime, _ := strconv.Atoi(strings.TrimRight(sets.Duration, "+"))

	//支援者ごとの支払い情報から入力条件を満たす支援者を判定
	for iUser, iPaySeqMap := range userPaySeqMap {
		m, _ := strconv.Atoi(sets.GetMonth)
		//ここら辺ちょっと細かく調べる↓
		for iYearMonth := checkTime; iYearMonth.Compare(checkTime.AddDate(0, -m+1, 0)) >= 0; iYearMonth = iYearMonth.AddDate(0, -1, 0) {
			if sets.Condition == "継続" {
				yearMonth := GetYearMonthFromTime(iYearMonth)

				if iPaySeqMap[yearMonth] == sets.Amount {
					counter = counter + 1
				} else {
					if strings.HasSuffix(sets.Duration, "+") {
						if counter >= durationTime {
							userResultMap[iUser] = true
						} else {
							userResultMap[iUser] = false
						}
					} else {
						if counter%durationTime == 0 {
							userResultMap[iUser] = true
						} else {
							userResultMap[iUser] = false
						}
					}
					break
				}
			} else if sets.Condition == "累積" {
				if iPaySeqMap[iUser] == sets.Amount {
					counter = counter + 1
				}
			}
		}

		if sets.Condition == "累積" {
			if strings.HasSuffix(sets.Duration, "+") {
				if counter >= durationTime {
					userResultMap[iUser] = true

				} else {
					userResultMap[iUser] = false

				}
			} else {
				if counter%durationTime == 0 {
					userResultMap[iUser] = true

				} else {
					userResultMap[iUser] = false

				}
			}
		}
	}

	fmt.Println(userResultMap)

	//フォーマットを開く
	f, err := excelize.OpenFile("入出力フォーマット.xlsx")
	if err != nil {
		panic("loadConfig excelize.OpenFile err:" + err.Error())
	}

	var outputSheetName string = "リスト"

	//リストの情報をクリアする代わりに一度リストシートを削除して新たにリストを作作成する
	f.DeleteSheet(outputSheetName)
	f.NewSheet(outputSheetName)

	var userColoumId int = 2
	var resultColoumId int = 3
	var userRowId int = 3
	var yearMonthColoumId int = 4
	var yearMonthRowId int = 2
	var firstIter bool = true

	//判定した情報をExcelに出力していく
	for iUser, iPaySeqMap := range userPaySeqMap {
		f.SetCellValue(outputSheetName, coordinatesToCellName(userColoumId, userRowId), iUser)

		if userResultMap[iUser] {
			f.SetCellValue(outputSheetName, coordinatesToCellName(resultColoumId, userRowId), "対象")
		}

		m, _ := strconv.Atoi(sets.GetMonth)

		//TODO細かい正しさは確かめる↓
		for iYearMonth := checkTime.AddDate(0, -m+1, 0); iYearMonth.Compare(checkTime) <= 0; iYearMonth = iYearMonth.AddDate(0, 1, 0) {
			yearMonth := GetYearMonthFromTime(iYearMonth)

			if firstIter {
				f.SetCellValue(outputSheetName, coordinatesToCellName(userColoumId, yearMonthRowId), "支援者名")
				f.SetCellValue(outputSheetName, coordinatesToCellName(resultColoumId, yearMonthRowId), "対象か？")
				f.SetCellValue(outputSheetName, coordinatesToCellName(yearMonthColoumId, yearMonthRowId), yearMonth)

			}
			if _, ok := iPaySeqMap[yearMonth]; ok {
				f.SetCellValue(outputSheetName, coordinatesToCellName(yearMonthColoumId, userRowId), iPaySeqMap[yearMonth])

			}
			yearMonthColoumId = yearMonthColoumId + 1

		}
		userRowId = userRowId + 1
		firstIter = false

	}
	f.Save()
	f.Close()

}

func coordinatesToCellName(columnId int, rowId int) string {
	ret, err := excelize.CoordinatesToCellName(columnId, rowId)
	if err != nil {
		panic("coordinatesToCellName err:" + err.Error())
	}
	return ret
}

func GetYearMonthFromTime(tm time.Time) string {
	return tm.Format("2006-01")
}
