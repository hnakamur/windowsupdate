package main

import (
	"fmt"

	wu "github.com/hnakamur/windowsupdate"
	"github.com/mattn/go-ole"
)

func installImportantUpdates() error {
	session, err := wu.NewSession()
	if err != nil {
		return err
	}
	defer session.Release()

	fmt.Printf("Start searching...\n")
	updatesToInstall, err := session.Search("IsInstalled=0 and Type='Software' and AutoSelectOnWebSites=1")
	if err != nil {
		return err
	}

	if len(updatesToInstall) == 0 {
		fmt.Printf("No important updates available. exiting\n")
		return nil
	}

	fmt.Printf("%d important updates available.\n", len(updatesToInstall))

	updatesToDownload := selectUpdatesToDownload(updatesToInstall)
	if len(updatesToDownload) > 0 {
		fmt.Printf("%d important updates to download\n", len(updatesToDownload))
		err = session.Download(updatesToDownload)
		if err != nil {
			return err
		}
		fmt.Printf("Downloaded %d important updates\n", len(updatesToDownload))
	} else {
		fmt.Printf("%d important updates are already downloaded\n", len(updatesToDownload))
	}

	result, err := session.Install(updatesToInstall)
	if err != nil {
		return err
	}

	fmt.Printf("Installed %d important updates. ResultCode=%d, RebootRequired=%v\n", len(updatesToInstall), result.ResultCode, result.RebootRequired)
	for i := 0; i < len(updatesToInstall); i++ {
		u := updatesToInstall[i]
		ur := result.UpdateResults[i]
		fmt.Printf("ID=%s, Title=%s, ResultCode=%d, RebootRequired=%v\n", u.ID, u.Title, ur.ResultCode, ur.RebootRequired)
	}

	return nil
}

func selectUpdatesToDownload(updatesToInstall []wu.Update) []wu.Update {
	updates := []wu.Update{}
	for _, update := range updatesToInstall {
		if !update.Downloaded {
			updates = append(updates, update)
		}
	}
	return updates
}

func main() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	err := installImportantUpdates()
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}
}
