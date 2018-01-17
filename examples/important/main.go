package main

import (
	"fmt"

	"github.com/hnakamur/windowsupdate"
	"github.com/go-ole/go-ole"
)

func main() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	updatesToInstall, result, err := windowsupdate.InstallImportantUpdates()
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}

	fmt.Printf("Installed %d important updates. ResultCode=%d, RebootRequired=%v\n", len(updatesToInstall), result.ResultCode, result.RebootRequired)
	for i := 0; i < len(updatesToInstall); i++ {
		u := updatesToInstall[i]
		ur := result.UpdateResults[i]
		fmt.Printf("ID=%s, Title=%s, ResultCode=%d, RebootRequired=%v\n", u.ID, u.Title, ur.ResultCode, ur.RebootRequired)
	}
}
