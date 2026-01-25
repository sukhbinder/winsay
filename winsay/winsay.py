from __future__ import print_function

import argparse
import win32com.client as wincl

speaker = wincl.Dispatch("SAPI.SpVoice")

def say(text):
    try:
        speaker.Speak(text)
    except Exception as ex:
        print("Error in speaking: {} with this error{} ".format(text, ex))


def main():
    """
    Say in windows
    """
    parser = create_parser()

    args = parser.parse_args()
    volume = args.volume
    if volume < 0:
        volume = 0
    if volume > 100:
        volume = 100
    speaker.Volume = volume

    if args.voice:
        try:
            voices = list(speaker.GetVoices())
            voice_map = {voice.GetDescription(): voice for voice in voices}
            if args.voice in voice_map:
                speaker.Voice = voice_map[args.voice]
            else:
                print(f"Warning: Voice '{args.voice}' not found. Using default voice.")
        except Exception as e:
            print(f"Warning: Could not set voice. Using default voice. Error: {e}")

    if args.input:
        content = open(args.input).read()
        say(content)

    if args.text:
        sentence = " ".join(args.text)
        say(sentence)

def create_parser():
    parser = argparse.ArgumentParser(description="Say in windows",
                                     formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument("text", type=str, nargs="*",
                        help="sentence to speak", default="")
    parser.add_argument("-i", "--input", type=str,
                        help="Text File to speak", default=None)
    parser.add_argument("-v", "--volume", type=int, default=100,
                        help="Volume (0-100)")

    try:
        voices = [voice.GetDescription() for voice in speaker.GetVoices()]
        voice_help = "Select a voice. Available voices:\n- " + "\n- ".join(voices)
    except Exception:
        voices = []
        voice_help = "Select a voice."

    parser.add_argument("--voice", type=str, default=None, help=voice_help)

    return parser


if __name__ == "__main__":
    main()
