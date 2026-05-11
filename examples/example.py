# SPDX-License-Identifier: 0BSD

import sys
import time
from datetime import datetime
from enum import Enum
from pathlib import Path

import schedule

from ahkunwrapped import Script, AhkExitException

choice = None
HOTKEY_SEND_CHOICE = 'F2'


class Event(Enum):
    QUIT, SEND_CHOICE, CLEAR_CHOICE, CHOOSE_MONTH, CHOOSE_DAY = range(5)


# `format_dict=` so we can use `{{VARIABLE}}` within example.ahk
ahk = Script.from_file(Path(__file__).parent / 'example.ahk', format_dict=globals())


def main() -> None:
    print("Scroll your mousewheel up and down in Notepad.")
    schedule.every(10).seconds.do(print_time)

    try:
        while True:
            # ahk.poll()  # detect exit, but all `ahk.` functions include this

            event = ahk.get('event')  # contains `ahk.poll()`
            if event:
                ahk.set('event', '')
                on_event(event)

            schedule.run_pending()
            time.sleep(0.1)
    except AhkExitException as e:
        sys.exit(e.args[0])


def print_time() -> None:
    print(f"It is now {datetime.now().time()}")


def on_event(event: str) -> None:
    global choice

    def get_choice() -> str:
        return choice or datetime.now().strftime('%#I:%M %p')

    if event == str(Event.QUIT):
        ahk.exit()
    if event == str(Event.CLEAR_CHOICE):
        choice = None
    if event == str(Event.SEND_CHOICE):
        ahk.call('Send', f'{get_choice()} ')
    if event == str(Event.CHOOSE_MONTH):
        choice = datetime.now().strftime('%b')
        ahk.call('Notify', f"Month is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.")
    if event == str(Event.CHOOSE_DAY):
        choice = datetime.now().strftime('%#d')
        ahk.call('Notify', f"Day is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.")


if __name__ == '__main__':
    main()
