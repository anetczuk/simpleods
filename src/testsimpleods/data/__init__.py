"""Test data and samples for unit tests."""

#
# Copyright (c) 2026, Arkadiusz Netczuk <dev.arnet@gmail.com>
# All rights reserved.
#
# This source code is licensed under the BSD 3-Clause license found in the
# LICENSE file in the root directory of this source tree.
#

import os

_SCRIPT_DIR = os.path.dirname(__file__)


def get_data_root_path() -> str:
    """Get path to data directory."""
    return _SCRIPT_DIR


def get_data_path(file_name: str) -> str:
    """Get path to file inside data directory."""
    return os.path.join(_SCRIPT_DIR, file_name)


def read_data(file_name: str) -> str:
    """Read file from data directory."""
    file_path = get_data_path(file_name)
    with open(file_path, encoding="utf-8") as file:
        return file.read()
