package windowsupdate

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func (s *Session) Download(updates []Update) error {
	downloader, err := toIDispatchErr(oleutil.CallMethod((*ole.IDispatch)(s), "CreateUpdateDownloader"))
	if err != nil {
		return err
	}

	coll, err := toUpdateCollection(updates)
	if err != nil {
		return err
	}
	_, err = oleutil.PutProperty(downloader, "Updates", coll)
	if err != nil {
		return err
	}

	_, err = toIDispatchErr(oleutil.CallMethod(downloader, "Download"))
	return err
}

func toUpdateCollection(updates []Update) (*ole.IDispatch, error) {
	unknown, err := oleutil.CreateObject("Microsoft.Update.UpdateColl")
	if err != nil {
		return nil, err
	}
	coll, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	for _, update := range updates {
		_, err := oleutil.CallMethod(coll, "Add", update.disp)
		if err != nil {
			return nil, err
		}
	}
	return coll, nil
}
