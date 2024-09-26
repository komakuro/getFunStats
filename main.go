package main

import (
	"encoding/json"
	"fmt"
	"os"
	"time"

	"github.com/sclevine/agouti"
)

type settings struct {
	LoginId  string `json:'loginId'`
	Password string `json:'password'`
}

func loadConfig() settings {
	f, err := os.Open("settings.json")
	if err != nil {
		panic("loadConfig os.Open err:" + err.Error())
	}
	defer f.Close()

	var cfg settings
	_ = json.NewDecoder(f).Decode(&cfg)

	return cfg
}

func main() {
	sets := loadConfig()

	// chromeを起動
	driver := agouti.ChromeDriver()
	driver.Start()
	defer driver.Stop() // chromeを終了

	page, _ := driver.NewPage()
	page.Navigate("https://hogehogefuga.fanbox.cc/manage/relationships")
	page.Screenshot("Screenshot01.png")

	loginIdForm := page.AllByXPath("//*[@id=\"app-mount-point\"]/div/div/div[4]/div[1]/div[2]/div/div/div/form/fieldset[1]/label/input")
	count, _ := loginIdForm.Count()
	fmt.Println("count", count)

	loginIdForm.Fill(sets.LoginId)

	passwordForm := page.AllByXPath("//*[@id=\"app-mount-point\"]/div/div/div[4]/div[1]/div[2]/div/div/div/form/fieldset[2]/label/input")
	pasCount, _ := passwordForm.Count()
	fmt.Println("pasCount", pasCount)

	passwordForm.Fill(sets.Password)

	page.Screenshot("Screenshot02.png")

	loginSubmit := page.AllByXPath("//*[@id=\"app-mount-point\"]/div/div/div[4]/div[1]/div[2]/div/div/div/form/button[1]")
	loginSubmitCount, _ := loginSubmit.Count()
	fmt.Println("loginSubmitCount", loginSubmitCount)

	loginSubmit.Submit()

	time.Sleep(3 * time.Second)

	page.Screenshot("Screenshot03.png")

	time.Sleep(5 * time.Second)
}
