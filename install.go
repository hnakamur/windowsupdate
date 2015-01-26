package windowsupdate

import (
	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

const (
	OrcNotStarted          = 0
	OrcInProgress          = 1
	OrcSucceeded           = 2
	OrcSucceededWithErrors = 3
	OrcFailed              = 4
	OrcAborted             = 5
)

type InstallationResult struct {
	disp           *ole.IDispatch
	RebootRequired bool
	ResultCode     int
	UpdateResults  []UpdateResult
}

type UpdateResult struct {
	RebootRequired bool
	ResultCode     int
}

func (s *Session) Install(updates []*Update) (*InstallationResult, error) {
	installer, err := toIDispatchErr(oleutil.CallMethod((*ole.IDispatch)(s), "CreateUpdateInstaller"))
	if err != nil {
		return nil, err
	}

	coll, err := toUpdateCollection(updates)
	if err != nil {
		return nil, err
	}
	_, err = oleutil.PutProperty(installer, "Updates", coll)
	if err != nil {
		return nil, err
	}

	resultDisp, err := toIDispatchErr(oleutil.CallMethod(installer, "Install"))
	if err != nil {
		return nil, err
	}

	return toInstallationResult(resultDisp, len(updates))
}

func toInstallationResult(resultDisp *ole.IDispatch, updateCount int) (*InstallationResult, error) {
	result := InstallationResult{disp: resultDisp}
	rebootRequired, err := toBoolErr(oleutil.GetProperty(resultDisp, "RebootRequired"))
	if err != nil {
		return nil, err
	}
	result.RebootRequired = rebootRequired

	resultCode, err := toInt64Err(oleutil.GetProperty(resultDisp, "ResultCode"))
	if err != nil {
		return nil, err
	}
	result.ResultCode = int(resultCode)

	for i := 0; i < updateCount; i++ {
		urDisp, err := toIDispatchErr(oleutil.CallMethod(resultDisp, "GetUpdateResult", i))
		if err != nil {
			return nil, err
		}
		rebootRequired, err := toBoolErr(oleutil.GetProperty(urDisp, "RebootRequired"))
		if err != nil {
			return nil, err
		}

		resultCode, err := toInt64Err(oleutil.GetProperty(urDisp, "ResultCode"))
		if err != nil {
			return nil, err
		}
		result.UpdateResults = append(result.UpdateResults,
			UpdateResult{ResultCode: int(resultCode), RebootRequired: rebootRequired})
	}
	return &result, nil
}
