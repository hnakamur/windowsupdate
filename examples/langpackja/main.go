package main

import (
	"fmt"

	"github.com/hnakamur/windowsupdate"
	"github.com/mattn/go-ole"
)

const (
	JapaneseLanguagePackUpdateID = "00a156d4-3876-4cd5-bd38-517679c6ba59"
)

func InstallJapaneseLanguagePack() error {
	session, err := windowsupdate.NewSession()
	if err != nil {
		return err
	}
	defer session.Release()

	fmt.Printf("Start searching...\n")
	update, err := session.FindByUpdateID(JapaneseLanguagePackUpdateID)
	if err != nil {
		return err
	}

	installed, err := update.Installed()
	if err != nil {
		return err
	}

	if installed {
		fmt.Printf("already installed. exiting\n")
		return nil
	}

	updates := []*windowsupdate.Update{update}

	downloaded, err := update.Downloaded()
	if downloaded {
		fmt.Printf("already downloaded, skip downloading\n")
	} else {
		err = session.Download(updates)
		if err != nil {
			return err
		}
		downloaded, err := update.Downloaded()
		if err != nil {
			return err
		}
		fmt.Printf("downloaded=%v\n", downloaded)
	}

	result, err := session.Install(updates)
	if err != nil {
		return err
	}

	fmt.Printf("ResultCode=%d, RebootRequired=%v\n", result.ResultCode, result.RebootRequired)
	for i, ur := range result.UpdateResults {
		fmt.Printf("UpdateResult[%d] ResultCode=%d, RebootRequired=%v\n", i, ur.ResultCode, ur.RebootRequired)
	}

	return nil
}

func main() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	err := InstallJapaneseLanguagePack()
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}
}
