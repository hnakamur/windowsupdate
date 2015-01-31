package main

import (
	"fmt"

	"github.com/hnakamur/windowsupdate"
	"github.com/mattn/go-ole"
)

func main() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	_, _, err := windowsupdate.InstallLanguagePack(windowsupdate.JapaneseLanguagePackUpdateID)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}
}
