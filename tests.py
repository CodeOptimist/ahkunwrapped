import itertools
import math
import random
import sys
import timeit
import warnings
from functools import partial
from inspect import currentframe, getframeinfo
from pathlib import Path

import hypothesis.strategies as st
import pytest
from hypothesis import given

import autohotkey
from autohotkey import Script

ahk = Script.from_file(Path('tests.ahk'))

# setup = r'''
# from autohotkey import Script
# ahk = Script('Echo(val) {\nreturn % val\n}')
# '''
# print(timeit.timeit("ahk.f('Echo', [' '] * 5000)", setup=setup, number=100))
# print(timeit.timeit("ahk.f_main('Echo', [' '] * 5000)", setup=setup, number=100))


def test_utf16_internals():
    assert ahk.f('HasUtf16Internals')


@given(st.sampled_from([ahk.f, ahk.f_main]))
def test_smile(func):
    assert func('GetSmile') == '🙂'


@given(st.sampled_from([ahk.call, ahk.call_main, ahk.f, ahk.f_main]))
def test_missing_func(func):
    with pytest.raises(autohotkey.AhkFuncNotFoundError):
        func('BadFunc')


def test_user_exception():
    try:
        ahk.call('UserException')
    except autohotkey.AhkUserException as e:
        assert e.message == "UserException"
        assert e.what == "example what"
        assert e.extra == "example extra"
        assert e.file == ahk.file
        line_num = 1 + next(num for (num, line) in enumerate(ahk.script.split('\n')) if line.startswith('    throw Exception("UserException"'))
        assert e.line == line_num


# if fail, adjust its stacklevel=
def test_warning_lineno():
    with warnings.catch_warnings(record=True) as w:
        ahk.call('_Py_StdErr', autohotkey.AhkWarning.__name__, "some generic warning")  # get directly because unused atm
        assert w[0].filename == getframeinfo(currentframe()).filename and w[0].lineno == currentframe().f_lineno - 1
        # eat the redundant output from call() finishing
        ahk.popen.stdout.readline()
        ahk.popen.stderr.readline()


# if fail, adjust its stacklevel=
def test_precision_warning_lineno():
    with warnings.catch_warnings(record=True) as w:
        echo(1 / 3)  # AhkLossOfPrecisionWarning
        assert w[0].filename == getframeinfo(currentframe()).filename and w[0].lineno == currentframe().f_lineno - 1


echo = partial(ahk.f, 'Echo')
echo_main = partial(ahk.f_main, 'Echo')


def set_get(val):
    ahk.set('myVar', val)
    return ahk.get('myVar')


result_funcs = st.sampled_from([echo, echo_main, set_get])


@given(result_funcs, st.booleans())
def test_bool(func, bool_):
    assert func(bool_) == bool_


@given(result_funcs, st.integers())
def test_int(func, int_):
    assert func(int_) == int_


@given(result_funcs, st.from_type(float))
def test_float(func, float_):
    if math.isnan(float_) or math.isinf(float_):
        with pytest.raises(autohotkey.AhkUnsupportedValueError):
            func(float_)
    else:
        ahk_float = float(f'{float_:.6f}')
        if ahk_float != float_:
            with pytest.warns(autohotkey.AhkLossOfPrecisionWarning):
                assert func(float_) == ahk_float
        else:
            assert func(float_) == float_


echo_raw = partial(ahk.f_raw, 'Echo')
echo_raw_main = partial(ahk.f_raw_main, 'Echo')


def set_get_raw(val):
    ahk.set('myVar', val)
    return ahk.get_raw('myVar')


raw_result_funcs = st.sampled_from([echo_raw, echo_raw_main, set_get_raw])
newlines = [''.join(x) for x in itertools.product('a\n\r', repeat=3)]


@given(raw_result_funcs, st.one_of(st.from_type(str), st.sampled_from(newlines)))
def test_str(func, str_):
    if '\0' in str_ or Script.SEPARATOR in str_:
        with pytest.raises(autohotkey.AhkUnsupportedValueError):
            func(str_)
    else:
        assert func(str_) == str_


@pytest.mark.filterwarnings('error')
@given(raw_result_funcs, st.text())
def test_text(func, text):
    try:
        assert func(text) == text
    except (autohotkey.AhkWarning, autohotkey.AhkUnsupportedValueError):
        return


@pytest.mark.filterwarnings("error")
@given(raw_result_funcs, st.text())
def test_long_text(func, text):
    try:
        assert func(text) == text
    except (autohotkey.AhkWarning, autohotkey.AhkUnsupportedValueError):
        return

    rand_len = random.randint(2000, 4000)
    # ahk.call('Copy', f"{repr(text)} * {rand_len}")
    long_text = text * rand_len
    # print(len(long_text), file=sys.stderr)
    assert func(long_text) == long_text
