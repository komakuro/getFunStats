package main

import (
	"archive/zip"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/widget"
	"golang.org/x/sys/windows/registry"

	"github.com/sclevine/agouti"

	"github.com/xuri/excelize/v2"
)

// 設定情報の構造体
type settings struct {
	LoginId    string
	Password   string
	CreatorId  string
	Duration   string
	Amount     string
	Condition  string
	GetMonth   string
	ChoiceFlag string
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

type newWindow struct {
	fyne.Window
}

func ReadCell(f *excelize.File, sheetName string, cellPosition string) string {
	ret, err := f.GetCellValue(sheetName, cellPosition)
	if err != nil {
		panic("readCell err:" + err.Error())
	}

	return ret
}

func loadConfig() config {
	f, err := os.Open("settings.json")
	if err != nil {
		panic("loadconfig os.Open err:" + err.Error())
	}
	defer f.Close()

	var cfg config
	_ = json.NewDecoder(f).Decode(&cfg)

	//fmt.Println(cfg)
	return cfg
}

func loadJson() settings {
	f, err := os.Open("./save.json")
	if err != nil {
		panic("loadconfig os.Open err:" + err.Error())
	}
	defer f.Close()

	var stg settings
	_ = json.NewDecoder(f).Decode(&stg)

	//fmt.Println(cfg)
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

func newhelpWindow(app fyne.App) *newWindow {

	app.Settings().SetTheme(&myTheme{})
	win := app.NewWindow("ヘルプ")

	mailText := widget.NewLabel("メールアドレス")
	mailText.TextStyle.Bold = true

	mailExplain := widget.NewLabel("FANBOXにログインする際のメールアドレスを入力してください")

	passText := widget.NewLabel("パスワード")
	passText.TextStyle.Bold = true

	passExplain := widget.NewLabel("FANBOXにログインする際のパスワードを入力してください")

	creatorText := widget.NewLabel("クリエイターID")
	creatorText.TextStyle.Bold = true

	creaTorExplain := widget.NewLabel("FANBOXで設定しているクリエイターIDを入力してください")

	durationText := widget.NewLabel("継続期間")
	durationText.TextStyle.Bold = true

	durationExplain := widget.NewLabel("対象となるプランの継続期間を入力してください\n継続期間ごとなら「半角数字」、継続期間以上なら「半角数字+」と入力してください")

	amountText := widget.NewLabel("継続プラン金額")
	amountText.TextStyle.Bold = true

	amountExplain := widget.NewLabel("対象となるプランの支援金額を入力してください\n指定金額だけが対象なら「半角数字」、指定金額以上が対象なら「半角数字+」と入力してください")

	conditionText := widget.NewLabel("継続可能条件")
	conditionText.TextStyle.Bold = true

	conditionExplain := widget.NewLabel("継続判定に必要な条件を「連続」「累積」の2種類から選んでください")

	monthText := widget.NewLabel("取得月数")
	monthText.TextStyle.Bold = true

	monthExplain := widget.NewLabel("現在の取得年月を基準として過去何か月分の情報を取得したいかを半角数字で入力してください")

	flagText := widget.NewLabel("当月で達成していなくても対象に含めるか")
	flagText.TextStyle.Bold = true

	flagExplain := widget.NewLabel("指定した継続期間の達成を、現在の取得年月より過去に達成した場合でも対象として含めるかどうかを\n「含める」「含めない」の2種類から選んでください")

	supButton := widget.NewButton("補足(継続期間と条件について)", func() { newSupWindow(app).Show() })
	closeButton := widget.NewButton("閉じる", func() { win.Close() })

	win.SetContent(container.NewVBox(
		mailText,
		mailExplain,
		passText,
		passExplain,
		creatorText,
		creaTorExplain,
		durationText,
		durationExplain,
		amountText,
		amountExplain,
		conditionText,
		conditionExplain,
		monthText,
		monthExplain,
		flagText,
		flagExplain,
		supButton,
		closeButton,
	))

	win.CenterOnScreen()

	return &newWindow{win}
}

func newSupWindow(app fyne.App) *newWindow {

	app.Settings().SetTheme(&myTheme{})
	win := app.NewWindow("補足")

	supTitle := widget.NewLabel("継続期間と条件の設定について")
	supTitle.TextStyle.Bold = true
	supText1 := widget.NewLabel("以下の支払い履歴だった支援者に対して\n支援金額1000で実行する場合を例にします")
	supText2 := widget.NewLabel("①2023/7に実行した場合")
	supText2.TextStyle.Bold = true
	supText3 := widget.NewLabel("継続期間:6 継続可能条件:連続の場合　⇒対象外（2023/2に支援していないため）\n継続期間:6 継続可能条件:累積の場合　⇒対象（実行した2024/7で累計6ヶ月をちょうど達成したため）")
	supText4 := widget.NewLabel("②2023/8に実行した場合")
	supText4.TextStyle.Bold = true
	supText5 := widget.NewLabel("継続期間:6 継続可能条件:連続の場合　⇒対象（連続で6ヶ月をちょうど達成したため）\n継続期間:6 継続可能条件:累積の場合　⇒対象外（2024/7で達成し、支援1ヶ月目の判定になるため）")

	var data = [][]string{
		{"2024/01", "2024/02", "2024/03", "2024/04", "2024/05", "2024/06", "2024/07", "2024/08"},
		{"1000", "", "1000", "1000", "1000", "1000", "1000", "1000"},
	}

	supTable := widget.NewTable(
		func() (int, int) {
			return len(data), len(data[0])
		},
		func() fyne.CanvasObject {
			return widget.NewLabel("wide content")
		},
		func(i widget.TableCellID, o fyne.CanvasObject) {
			o.(*widget.Label).SetText(data[i.Row][i.Col])
		})

	for i := 0; i < len(data[0]); i++ {
		supTable.SetColumnWidth(i, 75)
	}

	supTable.StickyRowCount = 1

	closeButton := widget.NewButton("閉じる", func() { win.Close() })

	win.SetContent(container.NewVBox(
		supTitle,
		supText1,
		supTable,
		supText2,
		supText3,
		supText4,
		supText5,
		closeButton,
	))

	win.CenterOnScreen()

	return &newWindow{win}

}

func newSaveWindow(app fyne.App) *newWindow {

	app.Settings().SetTheme(&myTheme{})
	win := app.NewWindow("保存完了")

	saveText := widget.NewLabel("保存が完了しました")
	saveText.TextStyle.Bold = true

	closeButton := widget.NewButton("閉じる", func() { win.Close() })

	win.SetContent(container.NewVBox(
		saveText,
		closeButton,
	))

	win.CenterOnScreen()

	return &newWindow{win}
}

func newCompWindow(count int, txt string, app fyne.App) *newWindow {

	app.Settings().SetTheme(&myTheme{})
	win := app.NewWindow("出力完了")

	var compText string

	if count == 0 {
		compText = "出力に成功しました"
	} else {
		compText = "出力に失敗しました\n詳細はErrorLog.txtを確認してください"

		var logFile = GetFile("ErrorLog.txt")
		WriteFile(logFile, txt)
	}

	compLabel := widget.NewLabel(compText)
	compLabel.TextStyle.Bold = true

	closeButton := widget.NewButton("閉じる", func() { win.Close() })

	win.SetContent(container.NewVBox(
		compLabel,
		closeButton,
	))

	win.CenterOnScreen()

	return &newWindow{win}
}

func getChromeDriver() {

	//レジストリからChromeのバージョンを取得
	k, err := registry.OpenKey(registry.CURRENT_USER, `Software\Google\Chrome\BLBeacon`, registry.QUERY_VALUE)
	if err != nil {
		log.Fatal(err)
	}
	defer k.Close()

	kStr, _, err := k.GetStringValue("version")
	if err != nil {
		log.Fatal(err)
	}

	//fmt.Println("Google Chrome version:", kStr)

	//一致するバージョンのChromeDriverをダウンロード
	url := "https://storage.googleapis.com/chrome-for-testing-public/" + kStr + "/win64/chromedriver-win64.zip"

	if err := DownloadFile("chromedriver-win64.zip", url); err != nil {
		panic(err)
	}

	//ダウンロードしたzipファイルを解凍する
	rootPath, _ := os.Getwd()
	rootDir := filepath.Dir(rootPath)

	zipPath := filepath.Join(rootDir, "scrapingFanbox", "chromedriver-win64.zip")
	destDir := filepath.Join(rootDir, "scrapingFanbox", "output")

	if err := unZip(zipPath, destDir); err != nil {
		panic(err)
	}

	//解凍したChromeDriverをコピーする
	exePath := filepath.Join(rootDir, "scrapingFanbox", "output", "chromedriver-win64", "chromedriver-win64", "chromedriver.exe")
	outDir := "C:\\ProgramData\\scrapingFanbox\\chromedriver.exe"

	copyFile(exePath, outDir)

	//ダウンロードしたzipファイルと解凍先ディレクトリを削除
	os.Remove("chromedriver-win64.zip")
	os.RemoveAll("output")

	//コピー先のディレクトリにパスを通す
	os.Setenv("PATH", "C:\\ProgramData\\scrapingFanbox")

}

// 指定のURLからファイルをダウンロードする
func DownloadFile(filepath string, url string) error {

	resp, err := http.Get(url)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	out, err := os.Create(filepath)
	if err != nil {
		return err
	}
	defer out.Close()

	_, err = io.Copy(out, resp.Body)
	return err
}

// unZip zipファイルを展開する
func unZip(src, dest string) error {
	r, err := zip.OpenReader(src)
	if err != nil {
		return err
	}
	defer r.Close()

	ext := filepath.Ext(src)
	rep := regexp.MustCompile(ext + "$")
	dir := filepath.Base(rep.ReplaceAllString(src, ""))

	destDir := filepath.Join(dest, dir)
	// ファイル名のディレクトリを作成する
	if err := os.MkdirAll(destDir, os.ModeDir); err != nil {
		return err
	}

	for _, f := range r.File {
		if f.Mode().IsDir() {
			// ディレクトリは無視して構わない
			continue
		}
		if err := saveUnZipFile(destDir, *f); err != nil {
			return err
		}
	}

	return nil
}

// saveUnZipFile 展開したZipファイルをそのままローカルに保存する
func saveUnZipFile(destDir string, f zip.File) error {
	// 展開先のパスを設定する
	destPath := filepath.Join(destDir, f.Name)
	// 子孫ディレクトリがあれば作成する
	if err := os.MkdirAll(filepath.Dir(destPath), f.Mode()); err != nil {
		return err
	}
	// Zipファイルを開く
	rc, err := f.Open()
	if err != nil {
		return err
	}
	defer rc.Close()
	// 展開先ファイルを作成する
	destFile, err := os.Create(destPath)
	if err != nil {
		return err
	}
	defer destFile.Close()
	// 展開先ファイルに書き込む
	if _, err := io.Copy(destFile, rc); err != nil {
		return err
	}

	return nil
}

func copyFile(srcDir string, dstDir string) {

	src, err := os.Open(srcDir)
	if err != nil {
		panic(err)
	}
	defer src.Close()

	dst, err := os.Create(dstDir)
	if err != nil {
		panic(err)
	}
	defer dst.Close()

	_, err = io.Copy(dst, src)
	if err != nil {
		panic(err)
	}

}

func main() {

	//パラメータ設定の取得
	cfgs := loadConfig()

	var sets settings
	//保存データがあれば保存情報の取得
	if _, err := os.Stat("save.json"); err == nil {
		sets = loadJson()
	}

	mainApp := app.New()
	mainApp.Settings().SetTheme(&myTheme{})
	win := mainApp.NewWindow("scrapingFANBOX")

	initText := widget.NewLabel("初期設定")
	initText.TextStyle.Bold = true

	entry1 := widget.NewEntry()
	entry2 := widget.NewEntry()
	entry3 := widget.NewEntry()

	initForm := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "メールアドレス", Widget: entry1},
			{Text: "パスワード", Widget: entry2},
			{Text: "クリエイターID", Widget: entry3},
		},
	}

	setText := widget.NewLabel("継続条件設定")
	setText.TextStyle.Bold = true

	entry4 := widget.NewEntry()

	setForm1 := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "　　　継続期間", Widget: entry4},
		},
	}

	radioText := widget.NewLabel("　継続可能条件")
	radioText.TextStyle.Bold = true

	setRadio := widget.NewRadioGroup([]string{"連続", "累計"}, func(value string) {
		sets.Condition = value
	})

	entry5 := widget.NewEntry()
	entry6 := widget.NewEntry()

	setForm2 := &widget.Form{
		Items: []*widget.FormItem{ // we can specify items in the constructor
			{Text: "継続プラン金額", Widget: entry5},
			{Text: "取得月数", Widget: entry6},
		},
	}

	checkText := widget.NewLabel("過去の達成対象者を含めるか")
	checkText.TextStyle.Bold = true

	setCheck := widget.NewRadioGroup([]string{"含めない", "含める"}, func(value string) {
		sets.ChoiceFlag = value
	})

	helpButton := widget.NewButton("ヘルプ", func() { newhelpWindow(mainApp).Show() })

	saveButton := widget.NewButton("保存", func() {
		saveSet := settings{
			LoginId:    entry1.Text,
			Password:   entry2.Text,
			CreatorId:  entry3.Text,
			Duration:   entry4.Text,
			Amount:     entry5.Text,
			Condition:  setRadio.Selected,
			GetMonth:   entry6.Text,
			ChoiceFlag: setCheck.Selected,
		}

		s, err := json.MarshalIndent(&saveSet, "", "\t")
		if err != nil {
			fmt.Println("Error marshalling to JSON:", err)
			return
		}

		f := GetFile("save.json")

		f.Write(s)

		newSaveWindow(mainApp).Show()
	})
	bootButton := widget.NewButton("実行", func() {

		count, txt := bootScraping(sets, cfgs)

		newCompWindow(count, txt, mainApp).Show()

	})

	exitButton := widget.NewButton("終了", func() { mainApp.Quit() })

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
		helpButton,
		saveButton,
		bootButton,
		exitButton,
	))

	//保存設定があればフォームに設定
	if _, err := os.Stat("save.json"); err == nil {
		entry1.SetText(sets.LoginId)
		entry2.SetText(sets.Password)
		entry3.SetText(sets.CreatorId)
		entry4.SetText(sets.Duration)
		entry5.SetText(sets.Amount)
		entry6.SetText(sets.GetMonth)
		setRadio.SetSelected(sets.Condition)
		setCheck.SetSelected(sets.ChoiceFlag)
	}

	win.Resize(fyne.NewSize(600, 400))
	win.CenterOnScreen()
	win.ShowAndRun()

}

func bootScraping(sets settings, cfgs config) (int, string) {

	//エラー出力用の定義
	var errorCount int = 0
	var errorTxt string

	//ChromeDriverを取得
	getChromeDriver()

	// chromeを起動
	driver := agouti.ChromeDriver()
	driver.Start()
	defer driver.Stop() // chromeを終了

	//Cookieの格納されたディレクトリが存在するか確認
	userDir, _ := os.UserHomeDir()
	cookieDir := userDir + "\\AppData\\Local\\Google\\Chrome\\User Data\\Default"

	//支援者一覧ページを開く
	page, _ := driver.NewPage(
		agouti.Desired(agouti.Capabilities{
			"chromeOptions": map[string][]string{
				"args": {
					"user-data-dir=" + cookieDir,
					"--disable-gpu",
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
							if sets.ChoiceFlag == "含める" {
								if counter >= durationTime {
									userResultMap[iUser] = true
								} else {
									userResultMap[iUser] = false
								}
							} else {
								userResultMap[iUser] = false
							}
						} else if sets.ChoiceFlag == "含める" {
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
							if sets.ChoiceFlag == "含める" {
								if counter >= durationTime {
									userResultMap[iUser] = true
								} else {
									userResultMap[iUser] = false
								}
							} else {
								userResultMap[iUser] = false
							}
						} else if sets.ChoiceFlag == "含める" {
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

	//ログを出力
	//payStatsListの情報を出力
	listStr := ""
	for i := 0; i < len(payStatsList); i++ {
		listStr = listStr + payStatsList[i].UserName + ","
		listStr = listStr + payStatsList[i].PayTime + ","
		listStr = listStr + payStatsList[i].PayAmount + "," + "\n"
	}

	//userPaySeqMapの情報を出力
	payStr := ""
	for iUser, iPaySeqMap := range userPaySeqMap {

		for iYearMonth := range iPaySeqMap {
			payStr = payStr + iUser + ","
			payStr = payStr + iYearMonth + ","
			payStr = payStr + strconv.Itoa(iPaySeqMap[iYearMonth]) + "\n"
		}
	}

	//userResultMapの情報を出力
	resultStr := ""
	for iUser := range userResultMap {
		resultStr = resultStr + iUser + ","
		resultStr = resultStr + strconv.FormatBool(userResultMap[iUser]) + "\n"
	}

	var listLogFile = GetFile("LogList.txt")
	WriteFile(listLogFile, listStr)

	var payLogFile = GetFile("LogPay.txt")
	WriteFile(payLogFile, payStr)

	var resultLogFile = GetFile("LogResult.txt")
	WriteFile(resultLogFile, resultStr)

	return errorCount, errorTxt

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
