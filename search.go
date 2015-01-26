package windowsupdate

import (
	"errors"

	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

type Update struct {
	disp       *ole.IDispatch
	ID         string
	Title      string
	Downloaded bool
	Installed  bool
}

var UpdateNotFoundError = errors.New("Update not found")

func (s *Session) FindByUpdateID(updateID string) (Update, error) {
	updates, err := s.Search("UpdateID='" + updateID + "'")
	if err != nil {
		return Update{}, err
	}
	if len(updates) == 0 {
		return Update{}, UpdateNotFoundError
	}
	return updates[0], nil
}

func (s *Session) Search(criteria string) ([]Update, error) {
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

	return toUpdates(updatesDisp)
}

func toUpdates(updatesDisp *ole.IDispatch) ([]Update, error) {
	count, err := toInt64Err(oleutil.GetProperty(updatesDisp, "Count"))
	if err != nil {
		return nil, err
	}

	var updates []Update
	for i := 0; i < int(count); i++ {
		updateDisp, err := toIDispatchErr(oleutil.GetProperty(updatesDisp, "Item", i))
		if err != nil {
			return nil, err
		}

		update, err := toUpdate(updateDisp)
		if err != nil {
			return nil, err
		}

		updates = append(updates, update)
	}
	return updates, nil
}

func toUpdate(updateDisp *ole.IDispatch) (Update, error) {
	update := Update{disp: updateDisp}
	identity, err := toIDispatchErr(oleutil.GetProperty(updateDisp, "Identity"))
	if err != nil {
		return update, err
	}

	id, err := toStrErr(oleutil.GetProperty(identity, "UpdateID"))
	if err != nil {
		return update, err
	}
	update.ID = id

	title, err := toStrErr(oleutil.GetProperty(updateDisp, "Title"))
	if err != nil {
		return update, err
	}
	update.Title = title

	downloaded, err := toBoolErr(oleutil.GetProperty(updateDisp, "IsDownloaded"))
	if err != nil {
		return update, err
	}
	update.Downloaded = downloaded

	installed, err := toBoolErr(oleutil.GetProperty(updateDisp, "IsInstalled"))
	if err != nil {
		return update, err
	}
	update.Installed = installed

	return update, nil
}
