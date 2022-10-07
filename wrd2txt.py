#!/usr/bin/python3

import os
import sys

import docx


class App:
    def __init__(self):
        self._VERSION = "v1.0.0"
        self._HELP = f"""wrd2txt {self._VERSION}\nUsage:\n\t./wrd2txt.py (<option>|<input_file>)\nOptions:\n\t-h, --help\tShow this help message and exit\n\t-v, --version\tShow version and exit\n"""
        self._ERROR_PREFIX = "Error: "

    def show_help(self):
        print(self._HELP)

    def show_version(self):
        print("wrd2txt " + self._VERSION)

    def show_error(self, message: str):
        print(self._colorize(self._ERROR_PREFIX + message), file=sys.stderr)

    def _colorize(self, message: str, color: str = "red"):
        colors = {
            "red": "\033[31m",
            "reset": "\033[0m",
        }
        return colors[color] + message + colors["reset"]

    def get_text(self, file_path: str):
        doc = docx.Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs])
        return text


def main() -> None:
    app = App()
    args = sys.argv[1:]
    if len(args) == 0 or len(args) > 1:
        app.show_help()
        return
    if args[0].startswith("-"):
        if args[0] == "-v" or args[0] == "--version":
            app.show_version()
        elif args[0] == "-h" or args[0] == "--help":
            app.show_help()
        else:
            app.show_error("Unknown option: " + args[0])
    else:
        file_path = args[0]
        if not os.path.exists(file_path):
            app.show_error(f"File '{file_path}' does not exist.")
            return
        text = app.get_text(file_path)
        print(text)


if __name__ == "__main__":
    main()
