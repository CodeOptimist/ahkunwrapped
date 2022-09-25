# Copyright (C) 2019-2022  Christopher S. Galpin.  Licensed under AGPL-3.0-or-later.  See /NOTICE.
import itertools
import math
import random
# noinspection PyUnresolvedReferences
import sys
import timeit
import warnings
from datetime import timedelta
from functools import partial
from inspect import currentframe, getframeinfo
from pathlib import Path

import hypothesis.strategies as st
import pytest
from hypothesis import given, settings

import ahkunwrapped as autohotkey
from ahkunwrapped import Script

ahk = Script.from_file(Path('tests.ahk'))


def print_timings():
    setup = r'''
from ahkunwrapped import Script
ahk = Script('Echo(val) {\nreturn val\n}')
'''

    for number in (100, 1000):
        print(f'number={number}'.rjust(30), "1 buffer".rjust(20), "".rjust(20), "~100 buffers".rjust(20))
        for func in ('call', 'f', 'call_main', 'f_main'):
            single_buffer = timeit.timeit(f"ahk.{func}('Echo', ' ' * 2000)", setup=setup, number=number)
            many_buffers = timeit.timeit(f"ahk.{func}('Echo', ' ' * 200000)", setup=setup, number=number)
            print(f"{func}('Echo', ...)".rjust(30), f'{single_buffer:.4f}'.rjust(20), f'x {many_buffers / single_buffer:.1f} ='.rjust(20), f'{many_buffers:.4f}'.rjust(20))


if __name__ == '__main__':
    print_timings()


def test_utf16_internals():
    assert ahk.f('HasUtf16Internals')


@given(st.sampled_from([ahk.f, ahk.f_main]))
def test_smile(f):
    assert f('GetSmile') == '🙂'


@given(st.sampled_from([ahk.call, ahk.call_main, ahk.f, ahk.f_main]))
def test_missing_func(func):
    with pytest.raises(autohotkey.AhkFuncNotFoundError):
        func('ThisDoesntExist')


# This test may fail the first time after a computer restart.
@settings(deadline=timedelta(seconds=1))
@given(st.sampled_from([ahk.call, ahk.f]), st.sampled_from([ahk.call_main, ahk.f_main]))
def test_main_thread_required(func, func_main):
    with pytest.raises(autohotkey.AhkCantCallOutInInputSyncCallError):
        func('ComMsGraphCall')
    func_main('ComMsGraphCall')


@given(st.sampled_from([ahk.call, ahk.call_main]), st.sampled_from([ahk.f, ahk.f_main]))
def test_main_thread_not_required(call, f):
    call('ComFsoTempName')
    assert f('ComFsoTempName').endswith('.tmp')


def test_userexception():
    try:
        ahk.call('UserException')
        assert False
    except autohotkey.AhkUserException as e:
        assert e.message == "UserException"
        assert e.what == "example what"
        assert e.extra == "example extra"
        assert e.file == ahk.file


def test_userexception_lineno():
    try:
        ahk.call('UserException')
        assert False
    except autohotkey.AhkUserException as e:
        line_num = 1 + next(num for (num, line) in enumerate(ahk.script.split('\n')) if line.startswith('    throw Exception("UserException"'))
        assert e.line == line_num


# Documenting that we can't distinguish between Exception() with good data and a contrived object with bad. They're the same within AHK.
@pytest.mark.xfail(strict=True)  # expected fail
def test_userexception_lineno_for_contrived():
    try:
        ahk.call('ContrivedException')
        assert False
    except autohotkey.AhkUserException as e:
        line_num = 1 + next(num for (num, line) in enumerate(ahk.script.split('\n')) if line.startswith('    throw {Message: "ContrivedException"'))
        assert e.line == line_num


def test_nonexception_warning():
    for i in range(1, 4):
        with pytest.warns(autohotkey.AhkCaughtNonExceptionWarning):
            with pytest.raises(autohotkey.AhkUserException):
                ahk.call(f'NonException{i}')


# Documenting that we can't distinguish between Exception() with good data and a contrived object with bad. They're the same within AHK.
@pytest.mark.xfail(strict=True)  # expected fail
def test_nonexception_warning_for_contrived():
    with pytest.warns(autohotkey.AhkCaughtNonExceptionWarning):
        with pytest.raises(autohotkey.AhkUserException):
            ahk.call(f'ContrivedException')


# if fail, adjust its stacklevel=
def test_nonexception_warning_lineno():
    for i in range(1, 4):
        with warnings.catch_warnings(record=True) as w:
            with pytest.raises(autohotkey.AhkUserException):
                ahk.call(f'NonException{i}')
            assert w[0].filename == getframeinfo(currentframe()).filename and w[0].lineno == currentframe().f_lineno - 1


# if fail, adjust its stacklevel=
def test_warning_lineno():
    with warnings.catch_warnings(record=True) as w:
        ahk.call('_Py_StdErr', autohotkey.AhkWarning.__name__, "some generic warning")  # get directly because unused atm
        assert w[0].filename == getframeinfo(currentframe()).filename and w[0].lineno == currentframe().f_lineno - 1
        # eat the redundant output from call() finishing
        ahk.popen.stdout.readline()
        ahk.popen.stderr.readline()


# warning covered in test_float()
# if fail, adjust its stacklevel=
def test_precisionwarning_lineno():
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
def test_bool(f, bool_):
    assert f(bool_) == bool_


@given(result_funcs, st.integers())
def test_int(f, int_):
    assert f(int_) == int_


@given(result_funcs, st.from_type(float))
def test_float(f, float_):
    if math.isnan(float_) or math.isinf(float_):
        with pytest.raises(autohotkey.AhkUnsupportedValueError):
            f(float_)
    else:
        ahk_float = float(f'{float_:.6f}')
        if ahk_float != float_:
            with pytest.warns(autohotkey.AhkLossOfPrecisionWarning):
                assert f(float_) == ahk_float
        else:
            assert f(float_) == float_


echo_raw = partial(ahk.f_raw, 'Echo')
echo_raw_main = partial(ahk.f_raw_main, 'Echo')


def set_get_raw(val):
    ahk.set('myVar', val)
    return ahk.get_raw('myVar')


raw_result_funcs = st.sampled_from([echo_raw, echo_raw_main, set_get_raw])
newlines = [''.join(x) for x in itertools.product('a\n\r', repeat=3)]


@given(raw_result_funcs, st.one_of(st.from_type(str), st.sampled_from(newlines)))
def test_str(f, str_):
    if '\0' in str_ or Script.SEPARATOR in str_:
        with pytest.raises(autohotkey.AhkUnsupportedValueError):
            f(str_)
    else:
        assert f(str_) == str_


@pytest.mark.filterwarnings('error')
@given(raw_result_funcs, st.text())
def test_text(f, text):
    try:
        assert f(text) == text
    except (autohotkey.AhkWarning, autohotkey.AhkUnsupportedValueError):
        return


@pytest.mark.filterwarnings('error')
@given(raw_result_funcs, st.text())
def test_long_text(f, text):
    try:
        assert f(text) == text
    except (autohotkey.AhkWarning, autohotkey.AhkUnsupportedValueError):
        return

    rand_len = random.randint(2000, 4000)
    # ahk.call('Copy', f"{repr(text)} * {rand_len}")
    long_text = text * rand_len
    # print(len(long_text), file=sys.stderr)
    assert f(long_text) == long_text


# At > 100 Scripts:
# >       win32job.AssignProcessToJobObject(self.job, handle)
# E       pywintypes.error: (50, 'AssignProcessToJobObject', 'The request is not supported.')
def test_job_script_limit():
    for _ in range(101):
        Script()
