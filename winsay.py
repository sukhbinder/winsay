from __future__ import print_function
import time
import argparse
import datetime
import win32com.client as wincl




def say(text):
    try:
        speaker = wincl.Dispatch("SAPI.SpVoice")
        speaker.Speak(text)
    except Exception as ex:
        print("Error in speaking: ".format(ex.msg))

def main():
    """
    Say in windows
    """
    parser = argparse.ArgumentParser(description="Say in windows")
    parser.add_argument("text" ,type=str, nargs="*", help="sentence to speak", default="")
    args = parser.parse_args()
    if args.text:
        sentence = " ".join(args.text)
        say(sentence)
    




if __name__ == "__main__":
    main()
