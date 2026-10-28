import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import excel_cleaner as ec  # noqa: E402

SEP = ec.build_sep_regex({"space": True, "dash": True, "comma": False, "dot": False, "slash": False})


@pytest.mark.parametrize("src,expected", [
    ("5892-1041-0085-9016", "5892104100859016"),
    ("۵۸۹۲-۱۰۴۱-۰۰۸۵-۹۰۱۶", "5892104100859016"),
    ("5892 1041 0085 9016", "5892104100859016"),
    ("1\u200c2\u200c3", "123"),                 # نیم‌فاصله
    ("۰۰۱۲۳", "00123"),                          # صفر اول حفظ شود
    ("٠٠٧٨", "0078"),                            # ارقام عربی
    ("تهران ۱۲۳", "تهران 123"),                  # متن مختلط: فقط ارقام
])
def test_convert_text_changes(src, expected):
    assert ec.convert_text(src, True, True, SEP, False) == expected


@pytest.mark.parametrize("src", ["a-b", "12.5", "1403/05/12", "hello"])
def test_convert_text_unchanged(src):
    assert ec.convert_text(src, True, True, SEP, False) is None


def test_comma_option():
    with_comma = ec.build_sep_regex({"space": True, "dash": True, "comma": True, "dot": False, "slash": False})
    assert ec.convert_text("۱٬۵۰۰٬۰۰۰", True, True, with_comma, False) == "1500000"
    assert ec.convert_text("۱٬۵۰۰٬۰۰۰", True, True, SEP, False) == "1,500,000"


def test_as_number_rules():
    assert ec.convert_text("5892-1041", True, True, SEP, True) == 58921041
    assert ec.convert_text("0012-3", True, True, SEP, True) == "00123"                 # صفر اول → متن
    assert ec.convert_text("5892-1041-0085-9016", True, True, SEP, True) == "5892104100859016"  # ۱۶ رقم → متن


def test_parse_address():
    assert ec.parse_address("A2:A5", 10, 5) == [(1, 0, 5, 1)]
    assert ec.parse_address("C:D", 10, 5) == [(0, 2, 10, 4)]
    assert ec.parse_address("3:4", 10, 5) == [(2, 0, 4, 5)]
    assert ec.parse_address("B7", 10, 5) == [(6, 1, 7, 2)]
    assert ec.parse_address("۲:۳", 10, 5) == [(1, 0, 3, 5)]          # شماره ردیف فارسی
    assert ec.parse_address("Z1:Z5", 10, 5) == []                      # خارج از محدوده
    with pytest.raises(ValueError):
        ec.parse_address("???", 10, 5)
