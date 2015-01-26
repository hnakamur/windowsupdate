package windowsupdate

import (
	"errors"

	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

type Update ole.IDispatch

var UpdateNotFoundError = errors.New("Update not found")

func (s *Session) FindByUpdateID(updateID string) (*Update, error) {
	updates, err := s.Search("UpdateID='" + updateID + "'")
	if err != nil {
		return nil, err
	}
	if len(updates) == 0 {
		return nil, UpdateNotFoundError
	}
	return updates[0], nil
}

func (s *Session) Search(criteria string) ([]*Update, error) {
	searcher, err := toIDispatchErr(oleutil.CallMethod((*ole.IDispatch)(s), "CreateUpdateSearcher"))
	if err != nil {
		return nil, err
	}

	result, err := toIDispatchErr(oleutil.CallMethod(searcher, "Search", criteria))
	if err != nil {
		return nil, err
	}

	updatesDisp, err := toIDispatchErr(oleutil.GetProperty(result, "Updates"))
	if err != nil {
		return nil, err
	}

	count, err := toInt64Err(oleutil.GetProperty(updatesDisp, "Count"))
	if err != nil {
		return nil, err
	}

	var updates []*Update
	for i := 0; i < int(count); i++ {
		update, err := toIDispatchErr(oleutil.GetProperty(updatesDisp, "Item", i))
		if err != nil {
			return nil, err
		}

		updates = append(updates, (*Update)(update))
	}
	return updates, nil
}

func (u *Update) ID() (string, error) {
	identity, err := toIDispatchErr(oleutil.GetProperty((*ole.IDispatch)(u), "Identity"))
	if err != nil {
		return "", err
	}

	return toStrErr(oleutil.GetProperty(identity, "UpdateID"))
}

func (u *Update) Title() (string, error) {
	return toStrErr(oleutil.GetProperty((*ole.IDispatch)(u), "Title"))
}

func (u *Update) Downloaded() (bool, error) {
	return toBoolErr(oleutil.GetProperty((*ole.IDispatch)(u), "IsDownloaded"))
}

func (u *Update) Installed() (bool, error) {
	return toBoolErr(oleutil.GetProperty((*ole.IDispatch)(u), "IsInstalled"))
}

/*
type Update struct {
	ID        string
	Title     string
	Installed bool
}

func (s *Session) FindByUpdateID(updateID string) (Update, error) {
	updates, err := s.Search("UpdateID='" + updateID + "'")
	if err != nil {
		return Update{ID: updateID}, err
	}
	if len(updates) == 0 {
		return Update{ID: updateID}, UpdateNotFoundError
	}
	return updates[0], nil
}

func (s *Session) Search(criteria string) ([]Update, error) {
	searcher, err := toIDispatchErr(oleutil.CallMethod((*ole.IDispatch)(s), "CreateupdateSearcher"))
	if err != nil {
		return nil, err
	}

	result, err := toIDispatchErr(oleutil.CallMethod(searcher, "Search", criteria))
	if err != nil {
		return nil, err
	}

	updatesDisp, err := toIDispatchErr(oleutil.GetProperty(result, "Updates"))
	if err != nil {
		return nil, err
	}

	count, err := toInt64Err(oleutil.GetProperty(updatesDisp, "Count"))
	if err != nil {
		return nil, err
	}

	var updates []Update
	for i := 0; i < int(count); i++ {
		update, err := toIDispatchErr(oleutil.GetProperty(updatesDisp, "Item", i))
		if err != nil {
			return nil, err
		}

		title, err := toStrErr(oleutil.GetProperty(update, "Title"))
		if err != nil {
			return nil, err
		}

		identity, err := toIDispatchErr(oleutil.GetProperty(update, "Identity"))
		if err != nil {
			return nil, err
		}

		updateID, err := toStrErr(oleutil.GetProperty(identity, "UpdateID"))
		if err != nil {
			return nil, err
		}

		fmt.Printf("title=%s, updateID=%s\n", title, updateID)

		installed, err := toBoolErr(oleutil.GetProperty(update, "IsInstalled"))
		if err != nil {
			return nil, err
		}

		updates = append(updates,
			Update{Title: title, ID: updateID, Installed: installed})
	}
	return updates, nil
}
*/

/*
type Searcher ole.IDispatch
type Result ole.IDispatch

func (s *Session) CreateSearcher() (*Searcher, error) {
	disp, err := toIDispatchErr(oleutil.CallMethod((*ole.IDispatch)(s), "CreateupdateSearcher"))
	if err != nil {
		return nil, err
	}
	return (*Searcher)(disp), nil
}

func (s *Searcher) Search(criteria string) (*Result, error) {
	disp, err := toIDispatchErr(oleutil.CallMethod((*ole.IDispatch)(s), "Search", criteria))
	if err != nil {
		return nil, err
	}
	return (*Result)(disp), nil
}

func (r *Result) ResultCode() (int, error) {
	code, err := toInt64Err(oleutil.GetProperty((*ole.IDispatch)(r), "ResultCode"))
	if err != nil {
		return 0, err
	}
	return int(code), nil
}
*/

/*
func (s *Session) IsUpdateInstalled(updateId string) (bool, error) {
}

func printInstalledUpdates() error {
	unknown, err := oleutil.CreateObject("Microsoft.Update.Session")
	if err != nil {
		return err
	}
	session, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}

	searcher, err := toIDispatchErr(oleutil.CallMethod(session, "CreateupdateSearcher"))
	if err != nil {
		return err
	}

	fmt.Println("Searching installed windows updates...")
	result, err := toIDispatchErr(oleutil.CallMethod(searcher, "Search", "IsInstalled=0 and Type='Software'"))
	if err != nil {
		return err
	}

	updates, err := toIDispatchErr(oleutil.GetProperty(result, "Updates"))
	if err != nil {
		return err
	}

	count, err := toInt64Err(oleutil.GetProperty(updates, "Count"))
	if err != nil {
		return err
	}

	fmt.Printf("Installed updates count is %d\n", count)
	for i := 0; i < int(count); i++ {
		update, err := toIDispatchErr(oleutil.GetProperty(updates, "Item", i))
		if err != nil {
			return err
		}

		title, err := toStrErr(oleutil.GetProperty(update, "Title"))
		if err != nil {
			return err
		}

		identity, err := toIDispatchErr(oleutil.GetProperty(update, "Identity"))
		if err != nil {
			return err
		}

		updateId, err := toStrErr(oleutil.GetProperty(identity, "UpdateID"))
		if err != nil {
			return err
		}

		fmt.Printf("%d: %s: %s\n", i, updateId, title)
	}

	return nil
}
func main() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	err := printInstalledUpdates()
	if err != nil {
		panic(err)
	}
}
*/
