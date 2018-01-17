package windowsupdate

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
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

func (s *Session) Install(updates []Update) (InstallationResult, error) {
	empty := InstallationResult{}
	installer, err := toIDispatchErr(oleutil.CallMethod((*ole.IDispatch)(s), "CreateUpdateInstaller"))
	if err != nil {
		return empty, err
	}

	coll, err := toUpdateCollection(updates)
	if err != nil {
		return empty, err
	}
	_, err = oleutil.PutProperty(installer, "Updates", coll)
	if err != nil {
		return empty, err
	}

	resultDisp, err := toIDispatchErr(oleutil.CallMethod(installer, "Install"))
	if err != nil {
		return empty, err
	}

	return toInstallationResult(resultDisp, len(updates))
}

func toInstallationResult(resultDisp *ole.IDispatch, updateCount int) (InstallationResult, error) {
	result := InstallationResult{disp: resultDisp}
	rebootRequired, err := toBoolErr(oleutil.GetProperty(resultDisp, "RebootRequired"))
	if err != nil {
		return result, err
	}
	result.RebootRequired = rebootRequired

	resultCode, err := toInt32Err(oleutil.GetProperty(resultDisp, "ResultCode"))
	if err != nil {
		return result, err
	}
	result.ResultCode = int(resultCode)

	for i := 0; i < updateCount; i++ {
		urDisp, err := toIDispatchErr(oleutil.CallMethod(resultDisp, "GetUpdateResult", i))
		if err != nil {
			return result, err
		}
		rebootRequired, err := toBoolErr(oleutil.GetProperty(urDisp, "RebootRequired"))
		if err != nil {
			return result, err
		}

		resultCode, err := toInt32Err(oleutil.GetProperty(urDisp, "ResultCode"))
		if err != nil {
			return result, err
		}
		result.UpdateResults = append(result.UpdateResults,
			UpdateResult{ResultCode: int(resultCode), RebootRequired: rebootRequired})
	}
	return result, nil
}
