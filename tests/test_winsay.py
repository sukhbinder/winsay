import sys
from unittest.mock import MagicMock

# Mock the win32com module before it is imported by winsay
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()

from winsay.winsay import create_parser


def test_create_parser():
    parser = create_parser()
    args = parser.parse_args(["hello"])
    assert args.text == ["hello"]
    assert args.input is None
    assert args.volume == 100  # Check default


def test_create_parser_volume():
    parser = create_parser()
    args = parser.parse_args(["--volume", "50"])
    assert args.volume == 50

    args = parser.parse_args(["-v", "75"])
    assert args.volume == 75
