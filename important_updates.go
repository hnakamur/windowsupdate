package windowsupdate

func InstallImportantUpdates() (updatesToInstall []Update, result InstallationResult, err error) {
	session, err := NewSession()
	if err != nil {
		return
	}
	defer session.Release()

	updatesToInstall, err = session.Search("IsInstalled=0 and Type='Software' and AutoSelectOnWebSites=1")
	if err != nil {
		return
	}

	if len(updatesToInstall) == 0 {
		return
	}

	updatesToDownload := selectUpdatesToDownload(updatesToInstall)
	if len(updatesToDownload) > 0 {
		err = session.Download(updatesToDownload)
		if err != nil {
			return
		}
	}

	result, err = session.Install(updatesToInstall)
	return
}

func selectUpdatesToDownload(updatesToInstall []Update) []Update {
	updates := []Update{}
	for _, update := range updatesToInstall {
		if !update.Downloaded {
			updates = append(updates, update)
		}
	}
	return updates
}
