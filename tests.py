from main import *

def test_marg_excel_to_py_list():
    result = marg_excel_to_py_list(excel_path="temp_files/temp.xls")
    assert isinstance(result, list)

    try:
        marg_excel_to_py_list("some wrong path")
    except Exception:
        pass
    else:
        assert False, "No exception on wrong path"
