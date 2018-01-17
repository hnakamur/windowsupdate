package windowsupdate

import (
	"errors"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Update struct {
	disp       *ole.IDispatch
	Identity         IUpdateIdentity
	Title      string
	IsDownloaded bool
	IsInstalled  bool
}

type IUpdateIdentity struct { 
	RevisionNumber int32
	UpdateID string
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
	count, err := toInt32Err(oleutil.GetProperty(updatesDisp, "Count"))
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

func toUpdate(updateDisp *ole.IDispatch) (update Update, err error) {
	update.disp = updateDisp
	identity, err := toIDispatchErr(oleutil.GetProperty(updateDisp, "Identity"))
	if err != nil {
		return update, err
	}

	if update.Identity.RevisionNumber, err = toInt32Err(oleutil.GetProperty(identity, "RevisionNumber")); err != nil {
		return update, err
	}
	if update.Identity.UpdateID, err = toStrErr(oleutil.GetProperty(identity, "UpdateID")); err != nil {
		return update, err
	}

	if update.Title, err = toStrErr(oleutil.GetProperty(updateDisp, "Title"));  err != nil {
		return update, err
	}

	if update.IsDownloaded, err = toBoolErr(oleutil.GetProperty(updateDisp, "IsDownloaded"));  err != nil {
		return update, err
	}

	if update.IsInstalled, err = toBoolErr(oleutil.GetProperty(updateDisp, "IsInstalled")); err != nil {
		return update, err
	}

	return update, nil
}
