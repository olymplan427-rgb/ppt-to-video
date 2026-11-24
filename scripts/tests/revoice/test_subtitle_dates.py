# -*- coding: utf-8 -*-
from revoice.compose import _split_long


def test_date_not_split_across_lines():
    t = "제2회 전국 학력평가 시험은 10월 21일부터 공식 홈페이지에서 모집하고 있습니다."
    for ch in _split_long(t, 28):
        assert not ch.endswith("10월"), ch
