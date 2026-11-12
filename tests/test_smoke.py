"""Altyapı duman testi.

Amaç: pytest keşfinin, `src` yolunun ve paket importunun çalıştığını
doğrulamak. İş mantığı testi değildir.
"""

import procurement


def test_package_imports():
    assert procurement.__name__ == "procurement"


def test_package_has_version():
    assert isinstance(procurement.__version__, str)
    assert procurement.__version__.count(".") == 2


def test_stage_marker():
    assert procurement.__version__ == "0.1.0"
