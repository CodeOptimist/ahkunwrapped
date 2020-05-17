import itertools
import math
from functools import partial

import hypothesis.strategies as st
import pytest
from hypothesis import assume, given

import autohotkey
from autohotkey import Script

ahk = Script.from_file('tests.ahk')
echo = partial(ahk.f, 'Echo')
echo_main = partial(ahk.f_main, 'Echo')


@given(st.sampled_from([ahk.f, ahk.f_main]))
def test_smile(func):
    assert func('GetSmile') == '🙂'


@given(st.sampled_from([ahk.call, ahk.call_main, ahk.f, ahk.f_main]))
def test_missing_func(func):
    with pytest.raises(autohotkey.AhkFuncNotFoundError):
        func('BadFunc')


def set_get(val, coerce_type=True):
    ahk.set('myVar', val)
    return ahk.get('myVar', coerce_type=coerce_type)


type_funcs = st.sampled_from([echo, echo_main, set_get])


@given(type_funcs, st.booleans())
def test_bool(func, bool_):
    assert func(bool_) == bool_


@given(type_funcs, st.integers())
def test_int(func, int_):
    assert func(int_) == int_


@given(type_funcs, st.from_type(float))
def test_float(func, float_):
    if math.isnan(float_) or math.isinf(float_):
        with pytest.raises(autohotkey.AhkUnsupportedValueError):
            func(float_)
    else:
        ahk_float = float(f'{float_:.6f}')
        assert func(float_) == ahk_float


newlines = [''.join(x) for x in itertools.product('a\n\r', repeat=3)]


@given(type_funcs, st.one_of(st.from_type(str), st.sampled_from(newlines)))
def test_str(func, str_):
    assume('\r' not in str_)
    if '\0' in str_ or Script.SEPARATOR in str_:
        with pytest.raises(autohotkey.AhkUnsupportedValueError):
            func(str_)
    else:
        assert func(str_, coerce_type=False) == str_
