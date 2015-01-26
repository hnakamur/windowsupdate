package windowsupdate

import (
	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

type Session ole.IDispatch

func NewSession() (*Session, error) {
	unknown, err := oleutil.CreateObject("Microsoft.Update.Session")
	if err != nil {
		return nil, err
	}
	disp, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	return (*Session)(disp), nil
}

func (s *Session) Release() {
	(*ole.IDispatch)(s).Release()
}
