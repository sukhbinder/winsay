import sys
from unittest.mock import MagicMock
import pytest

# Mock the win32com module before it is imported by winsay
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()

# Now we can import the module that depends on win32com
from winsay.winsay import create_parser, main

@pytest.fixture
def mock_speaker_and_voices(monkeypatch):
    """Mocks the global speaker object in winsay.winsay."""
    mock_voice1 = MagicMock()
    mock_voice1.GetDescription.return_value = "Voice1"
    mock_voice2 = MagicMock()
    mock_voice2.GetDescription.return_value = "Voice2"

    mock_speaker_obj = MagicMock()
    mock_speaker_obj.GetVoices.return_value = [mock_voice1, mock_voice2]

    monkeypatch.setattr("winsay.winsay.speaker", mock_speaker_obj)
    return mock_speaker_obj, mock_voice1, mock_voice2

def test_create_parser(mock_speaker_and_voices):
    parser = create_parser()
    args = parser.parse_args(["hello"])
    assert args.text == ["hello"]
    assert args.input is None
    assert args.volume == 100
    assert args.voice is None

def test_create_parser_volume(mock_speaker_and_voices):
    parser = create_parser()
    args = parser.parse_args(["--volume", "50"])
    assert args.volume == 50

    args = parser.parse_args(["-v", "75"])
    assert args.volume == 75

def test_create_parser_voice(mock_speaker_and_voices):
    parser = create_parser()
    args = parser.parse_args(["--voice", "Voice1"])
    assert args.voice == "Voice1"

def test_main_set_voice(monkeypatch, mock_speaker_and_voices):
    mock_speaker_obj, _, mock_voice2_obj = mock_speaker_and_voices
    monkeypatch.setattr(sys, "argv", ["winsay", "--voice", "Voice2", "some text"])
    main()
    assert mock_speaker_obj.Voice == mock_voice2_obj

def test_main_invalid_voice(monkeypatch, capsys, mock_speaker_and_voices):
    monkeypatch.setattr(sys, "argv", ["winsay", "--voice", "InvalidVoice", "some text"])
    main()
    captured = capsys.readouterr()
    assert "Warning: Voice 'InvalidVoice' not found. Using default voice." in captured.out

def test_main_volume_clamping(monkeypatch, mock_speaker_and_voices):
    mock_speaker_obj, _, _ = mock_speaker_and_voices

    # Test volume > 100
    monkeypatch.setattr(sys, "argv", ["winsay", "--volume", "150", "some text"])
    main()
    assert mock_speaker_obj.Volume == 100

    # Test volume < 0
    monkeypatch.setattr(sys, "argv", ["winsay", "--volume", "-10", "some text"])
    main()
    assert mock_speaker_obj.Volume == 0

    # Test volume in range
    monkeypatch.setattr(sys, "argv", ["winsay", "--volume", "50", "some text"])
    main()
    assert mock_speaker_obj.Volume == 50
