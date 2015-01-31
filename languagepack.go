package windowsupdate

const (
	JapaneseLanguagePackUpdateID = "00a156d4-3876-4cd5-bd38-517679c6ba59"
)

func InstallLanguagePack(languagePackID string) (updates []Update, result InstallationResult, err error) {
	session, err := NewSession()
	if err != nil {
		return
	}
	defer session.Release()

	update, err := session.FindByUpdateID(languagePackID)
	if err != nil {
		return
	}

	if update.Installed {
		return
	}

	updates = []Update{update}

	if !update.Downloaded {
		err = session.Download(updates)
		if err != nil {
			return
		}
	}

	result, err = session.Install(updates)
	return
}
