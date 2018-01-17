package windowsupdate

import "github.com/go-ole/go-ole"

func toIDispatchErr(result *ole.VARIANT, err error) (*ole.IDispatch, error) {
	if err != nil {
		return nil, err
	}
	return result.ToIDispatch(), nil
}

func toInt64Err(result *ole.VARIANT, err error) (int64, error) {
	if err != nil {
		return 0, err
	}
	return variantToInt64(result), nil
}

func toInt32Err(result *ole.VARIANT, err error) (int32, error) {
	if err != nil {
		return 0, err
	}
	return variantToInt32(result), nil
}

func toStrErr(result *ole.VARIANT, err error) (string, error) {
	if err != nil {
		return "", err
	}
	return variantToStr(result), nil
}

func toBoolErr(result *ole.VARIANT, err error) (bool, error) {
	if err != nil {
		return false, err
	}
	return variantToBool(result), nil
}

func variantToInt64(v *ole.VARIANT) int64 {
	return v.Value().(int64)
}

func variantToInt32(v *ole.VARIANT) int32 {
	return v.Value().(int32)
}

func variantToStr(v *ole.VARIANT) string {
	return v.Value().(string)
}

func variantToBool(v *ole.VARIANT) bool {
	return v.Value().(bool)
}
