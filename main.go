package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/widget"

	"github.com/sclevine/agouti"

	"github.com/xuri/excelize/v2"
)

// 設定情報の構造体
type settings struct {
	CreatorId  string
	LoginId    string
	Password   string
	pcUser     string
	Duration   string
	Amount     string
	Condition  string
	GetMonth   string
	choiceFlag string
}

// 支援者情報の構造体
type payStats struct {
	UserName  string
	PayTime   string
	PayAmount string
}

// 外部パラメータの構造体
type config struct {
	LoginWaitTime    int `json:"loginWaitTime"`
	InfoLoadWaitTime int `json:"infoLoadWaitTime"`
}

func readCell(f *excelize.File, sheetName string, cellPosition string) string {
	ret, err := f.GetCellValue(sheetName, cellPosition)
	if err != nil {
		panic("readCell err:" + err.Error())
	}

	return ret
}

func loadConfig() config {
	f, err := os.Open("./exe/settings.json")
	if err != nil {
		panic("loadconfig os.Open err:" + err.Error())
	}
	defer f.Close()

	var cfg config
	_ = json.NewDecoder(f).Decode(&cfg)

	//fmt.Println(cfg)
	return cfg
}

func loadSettings() settings {

	//フォーマットを開く
	f, err := excelize.OpenFile("入出力フォーマット.xlsx")
	if err != nil {
		panic("loadConfig excelize.OpenFile err:" + err.Error())
	}
	defer f.Close()

	var stg settings

	//フォーマット内の指定されたセルの値を取得
	stg.CreatorId = readCell(f, "設定", "C6")
	stg.LoginId = readCell(f, "設定", "C4")
	stg.Password = readCell(f, "設定", "C5")
	stg.pcUser = readCell(f, "設定", "C7")
	stg.Duration = readCell(f, "設定", "C11")
	stg.Amount = readCell(f, "設定", "C12")
	stg.Condition = readCell(f, "設定", "C13")
	stg.GetMonth = readCell(f, "設定", "C14")
	stg.choiceFlag = readCell(f, "設定", "C15")

	return stg
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

func UpdateTime(clock *widget.Label) {
	formatted := time.Now().Format("Time: 03:04:05")
	clock.SetText(formatted)
}

func main() {
	//設定の取得
	sets := loadSettings()

	//パラメータ設定の取得
	cfgs := loadConfig()

	//エラー出力用の定義
	var errorCount = 0
	var errorTxt string

	app := app.New()
	app.Settings().SetTheme(&myTheme{})
	win := app.NewWindow("scrapingFANBOX")

	initText := widget.NewLabel("初期設定")
	initText.TextStyle.Bold = true

	entry1 := widget.NewEntry()
	entry2 := widget.NewEntry()
	entry3 := widget.NewEntry()
	entry4 := widget.NewEntry()

	initForm := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "メールアドレス", Widget: entry1},
			{Text: "パスワード", Widget: entry2},
			{Text: "クリエイターID", Widget: entry3},
			{Text: "PCユーザー名", Widget: entry4},
		},
	}

	setText := widget.NewLabel("継続条件設定")
	setText.TextStyle.Bold = true

	entry5 := widget.NewEntry()

	setForm1 := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "　　　継続期間", Widget: entry5},
		},
	}

	radioText := widget.NewLabel("　継続可能条件")
	radioText.TextStyle.Bold = true

	setRadio := widget.NewRadioGroup([]string{"連続", "累計"}, func(value string) {
		log.Println("Radio set to", value)
	})

	entry6 := widget.NewEntry()
	entry7 := widget.NewEntry()
	entry8 := widget.NewEntry()

	setForm2 := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "継続プラン金額", Widget: entry6},
			{Text: "継続可能条件", Widget: entry7},
			{Text: "取得月数", Widget: entry8},
		},
	}

	checkText := widget.NewLabel("過去の達成対象者を含めるか")
	checkText.TextStyle.Bold = true

	setCheck := widget.NewRadioGroup([]string{"含めない", "含める"}, func(value string) {
		log.Println("Check set to", value)
	})

	saveButton := widget.NewButton("保存", func() { app.Quit() })
	bootButton := widget.NewButton("実行", func() { app.Quit() })

	win.SetContent(container.NewVBox(
		initText,
		initForm,
		setText,
		setForm1,
		container.NewHBox(
			radioText,
			setRadio,
		),
		setForm2,
		container.NewHBox(
			checkText,
			setCheck,
		),
		saveButton,
		bootButton,
	))

	win.Resize(fyne.NewSize(600, 400))
	win.CenterOnScreen()
	win.ShowAndRun()

	// chromeを起動
	driver := agouti.ChromeDriver()
	driver.Start()
	defer driver.Stop() // chromeを終了

	//Cookieの格納されたディレクトリが存在するか確認
	cookieDir := "C:\\Users\\" + sets.pcUser + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default"
	if f, err := os.Stat(cookieDir); os.IsNotExist(err) || !f.IsDir() {
		errorCount = errorCount + 1
		errorTxt = errorTxt + "指定のフォルダが見つかりません\nPCユーザー名が正しいか確認してください\n"
	}

	//支援者一覧ページを開く
	page, _ := driver.NewPage(
		agouti.Desired(agouti.Capabilities{
			"chromeOptions": map[string][]string{
				"args": {
					"user-data-dir=" + cookieDir,
				},
			},
		}),
	)
	page.Navigate("https://" + sets.CreatorId + ".fanbox.cc/manage/relationships")

	//ログインを行う
	fillForm := page.AllByClass("sc-bn9ph6-6")
	fillCount, _ := fillForm.Count()
	//fmt.Println("fillCount", fillCount)

	if fillCount != 0 {
		errorCount = errorCount + 1
		errorTxt = errorTxt + "支援者一覧ページにアクセスできません\nクリエイターID、ログインアドレス、ログインパスワードが正しいか確認してください\n確認コードの入力やreCapcha認証を求められた場合は\n一度GoogleChromeでFANBOXにログインをした上でもう一度実行してください\n"
	}

	for i := 0; i < fillCount; i++ {
		if i == 0 {
			fillForm.At(i).Fill(sets.LoginId)
		}
		if i == 1 {
			fillForm.At(i).Fill(sets.Password)
		}
	}

	loginSubmit := page.AllByClass("sc-2o1uwj-9")
	//loginSubmitCount, _ := loginSubmit.Count()
	//fmt.Println("loginSubmitCount", loginSubmitCount)

	loginSubmit.Submit()

	time.Sleep(time.Duration(cfgs.LoginWaitTime) * time.Second)

	//支援者一覧を取得
	supportUsers := page.AllByClass("Row__UserWrapper-sc-1xb9lq9-1")
	supportUsersCount, _ := supportUsers.Count()
	//fmt.Println("supportUsersCount", supportUsersCount)

	var payStatsList []payStats

	//各支援者ページに遷移して支払い日時と支払金額をスライスに格納していく
	for i := 0; i < supportUsersCount; i++ {
		supportUsers.At(i).Click()

		time.Sleep(1 * time.Second)

		records := page.AllByClass("SupportTransactionSection__Td-sc-17tc9du-3")
		recordsCount, _ := records.Count()
		title, _ := page.Title()
		//fmt.Println("user", title, "recordsCount", recordsCount)

		userName := strings.Split(title, "｜")[0]

		var oneLine payStats

		oneLine.UserName = userName

		for j := 0; j < recordsCount; j++ {
			record := records.At(j)
			txt, _ := record.Text()

			//fmt.Println("txt", txt)

			if j%2 == 0 {
				oneLine.PayTime = txt
			}

			if j%2 == 1 {
				oneLine.PayAmount = strings.ReplaceAll(txt, "\u00A5", "")
				//一行分の情報を取り終わったので、取得した支援者名、支払い日時、支払金額をスライスに格納
				payStatsList = append(payStatsList, oneLine)
				//次の一行の情報の取得処理を開始するにあたって、oneLineを初期化する
				oneLine = payStats{}
				oneLine.UserName = userName
			}

		}

		time.Sleep(time.Duration(cfgs.InfoLoadWaitTime) * time.Second)

		page.Back()

	}

	fmt.Println(payStatsList)

	//スライスの情報を整理するためのマップを作成
	var userPaySeqMap = make(map[string]map[string]int)

	//スライスに格納された内容の数だけマップに情報を格納
	for i := 0; i < len(payStatsList); i++ {

		var tmpPayUser string = payStatsList[i].UserName
		var tmpPaySeqMap = make(map[string]int)

		//マップ内に該当の支援者名が存在するか確認し、存在しなければ格納用のマップを作成
		if _, ok := userPaySeqMap[tmpPayUser]; ok {
			tmpPaySeqMap = userPaySeqMap[tmpPayUser]
		}

		var tmpPayDate string = payStatsList[i].PayTime
		tmpPayDate = tmpPayDate[:7]

		var tmpPayAmount string = payStatsList[i].PayAmount
		var tmpPayAmountInt, _ = strconv.Atoi(tmpPayAmount)
		var tmpPayAmountInt2, _ = tmpPaySeqMap[tmpPayDate]

		//マップ内に該当の支払い月が存在するか確認し、存在すれば支払金額を合算
		if _, ok := tmpPaySeqMap[tmpPayDate]; ok {
			var AmountSum int = tmpPayAmountInt2 + tmpPayAmountInt
			//var AmountSumStr string = strconv.Itoa(AmountSum)
			tmpPaySeqMap[tmpPayDate] = AmountSum

		} else {
			tmpPaySeqMap[tmpPayDate] = tmpPayAmountInt

		}

		userPaySeqMap[tmpPayUser] = tmpPaySeqMap
	}

	fmt.Println(userPaySeqMap)

	userResultMap := make(map[string]bool)
	checkTime := time.Now()
	checkMonth := GetYearMonthFromTime(checkTime)
	durationTime, _ := strconv.Atoi(strings.ReplaceAll(sets.Duration, "+", ""))
	amountInt, _ := strconv.Atoi(strings.ReplaceAll(sets.Amount, "+", ""))

	//支援者ごとの支払い情報から入力条件を満たす支援者を判定
	for iUser, iPaySeqMap := range userPaySeqMap {
		var counter int = 0
		m, _ := strconv.Atoi(sets.GetMonth)

		//現在の実行年月を基準として、ひと月ずつ取得月数分さかのぼっていく
		for iYearMonth := checkTime; iYearMonth.Compare(checkTime.AddDate(0, -m+1, 0)) >= 0; iYearMonth = iYearMonth.AddDate(0, -1, 0) {
			yearMonth := GetYearMonthFromTime(iYearMonth)
			payAmountInt := iPaySeqMap[yearMonth]

			if sets.Condition == "連続" {

				if strings.HasSuffix(sets.Amount, "+") {
					if payAmountInt >= amountInt {
						counter = counter + 1
					} else {
						if strings.HasSuffix(sets.Duration, "+") {
							if sets.choiceFlag == "含める" {
								if counter >= durationTime {
									userResultMap[iUser] = true
								} else {
									userResultMap[iUser] = false
								}
							} else {
								userResultMap[iUser] = false
							}
						} else if sets.choiceFlag == "含める" {
							if counter/durationTime > 0 {
								userResultMap[iUser] = true
							} else {
								userResultMap[iUser] = false
							}
						} else {
							if _, ok := iPaySeqMap[checkMonth]; ok && counter > 0 && counter%durationTime == 0 {
								userResultMap[iUser] = true
							} else {
								userResultMap[iUser] = false
							}
						}
						break
					}
				} else {
					if payAmountInt == amountInt {
						counter = counter + 1
					} else {
						if strings.HasSuffix(sets.Duration, "+") {
							if sets.choiceFlag == "含める" {
								if counter >= durationTime {
									userResultMap[iUser] = true
								} else {
									userResultMap[iUser] = false
								}
							} else {
								userResultMap[iUser] = false
							}
						} else if sets.choiceFlag == "含める" {
							if counter/durationTime > 0 {
								userResultMap[iUser] = true
							} else {
								userResultMap[iUser] = false
							}
						} else {
							if _, ok := iPaySeqMap[checkMonth]; ok && counter > 0 && counter%durationTime == 0 {
								userResultMap[iUser] = true
							} else {
								userResultMap[iUser] = false
							}
						}
						break
					}
				}
			} else if sets.Condition == "累積" {
				if strings.HasSuffix(sets.Amount, "+") {
					if payAmountInt >= amountInt {
						counter = counter + 1
					}
				} else {
					if payAmountInt == amountInt {
						counter = counter + 1
					}
				}
			}
		}

		if strings.HasSuffix(sets.Duration, "+") {
			if counter >= durationTime {
				userResultMap[iUser] = true
			} else {
				userResultMap[iUser] = false
			}
		} else {
			if counter > 0 && counter%durationTime == 0 {
				userResultMap[iUser] = true
			} else {
				userResultMap[iUser] = false
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
	err = f.DeleteSheet(outputSheetName)
	if err != nil {
		panic("loadConfig excelize.DeleteSheet err:" + err.Error())
	}
	_, err = f.NewSheet(outputSheetName)
	if err != nil {
		panic("loadConfig excelize.NewSheet err:" + err.Error())
	}
	err = f.SetColWidth(outputSheetName, "B", "B", 18)
	if err != nil {
		panic("loadConfig excelize.SetColWidth err:" + err.Error())
	}

	var userColoumId int = 2
	var resultColoumId int = 3
	var userRowId int = 3
	var yearMonthRowId int = 2
	var firstIter bool = true

	//判定した情報をExcelに出力していく
	for iUser, iPaySeqMap := range userPaySeqMap {

		var yearMonthColoumId int = 4

		userTitleCell, _ := excelize.CoordinatesToCellName(userColoumId, yearMonthRowId)
		resultTitleCell, _ := excelize.CoordinatesToCellName(resultColoumId, yearMonthRowId)
		userNameCell, _ := excelize.CoordinatesToCellName(userColoumId, userRowId)
		resultCell, _ := excelize.CoordinatesToCellName(resultColoumId, userRowId)

		f.SetCellValue(outputSheetName, userNameCell, iUser)

		if userResultMap[iUser] {
			f.SetCellValue(outputSheetName, resultCell, "対象")

			style, _ := f.NewStyle(&excelize.Style{
				Fill: excelize.Fill{Type: "pattern", Color: []string{"F4E511"}, Pattern: 1},
			})

			f.SetCellStyle(outputSheetName, resultCell, resultCell, style)
		}

		m, _ := strconv.Atoi(sets.GetMonth)

		//TODO細かい正しさは確かめる↓
		for iYearMonth := checkTime.AddDate(0, -m+1, 0); iYearMonth.Compare(checkTime) <= 0; iYearMonth = iYearMonth.AddDate(0, 1, 0) {
			yearMonth := GetYearMonthFromTime(iYearMonth)

			yaerMonthCell, _ := excelize.CoordinatesToCellName(yearMonthColoumId, yearMonthRowId)
			payAmountCell, _ := excelize.CoordinatesToCellName(yearMonthColoumId, userRowId)

			if firstIter {
				f.SetCellValue(outputSheetName, userTitleCell, "支援者名")
				f.SetCellValue(outputSheetName, resultTitleCell, "対象か？")
				f.SetCellValue(outputSheetName, yaerMonthCell, yearMonth)

			}
			if _, ok := iPaySeqMap[yearMonth]; ok {
				f.SetCellValue(outputSheetName, payAmountCell, iPaySeqMap[yearMonth])

			}
			yearMonthColoumId = yearMonthColoumId + 1

		}
		userRowId = userRowId + 1
		firstIter = false

	}
	err = f.Save()
	if err != nil {
		//panic("loadConfig excelize.Save err:" + err.Error())
		errorCount = errorCount + 1
		errorTxt = errorTxt + "出力情報を書き込めません\n入出力フォーマット.xlsxを閉じてからもう一度実行してください\n"
	}
	err = f.Close()
	if err != nil {
		panic("loadConfig excelize.Close err:" + err.Error())
	}

	//エラーがどこかで発生していた場合はエラー内容をログ出力する
	if errorCount == 0 {
		fmt.Println("出力に成功しました")

	} else {
		var logFile = GetFile("ErrorLog.txt")
		WriteFile(logFile, errorTxt)

		fmt.Println("出力に失敗しました、詳細はErrorLog.txtを確認してください")
	}

}

func CoordinatesToCellName(columnId int, rowId int) string {
	ret, err := excelize.CoordinatesToCellName(columnId, rowId)
	if err != nil {
		panic("coordinatesToCellName err:" + err.Error())
	}
	return ret
}

func GetYearMonthFromTime(tm time.Time) string {
	return tm.Format("2006-01")
}
