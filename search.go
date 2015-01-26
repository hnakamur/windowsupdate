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
