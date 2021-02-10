import sys
import time
from datetime import datetime
from enum import Enum
from pathlib import Path

from ahkunwrapped import Script, AhkExitException

choice = None
HOTKEY_SEND_CHOICE = 'F2'


class Event(Enum):
    QUIT, SEND_CHOICE, CLEAR_CHOICE, CHOOSE_MONTH, CHOOSE_DAY = range(5)


ahk = Script.from_file(Path('example.ahk'), format_dict=globals())


def main() -> None:
    ts = 0
    while True:
        exit_code = ahk.poll()
        if exit_code:
            sys.exit(exit_code)

        try:
            s_elapsed = time.time() - ts
            if s_elapsed >= 60:
                ts = time.time()
                print_minute()

            event = ahk.get('event')
            if event:
                ahk.set('event', '')
                on_event(event)
        except AhkExitException:
            print("Graceful exit.")
            sys.exit(0)
        time.sleep(0.01)


def print_minute() -> None:
    print(f'It is now {datetime.now().time()}')


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
        ahk.call('ToolTip', f'Month is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.')
    if event == str(Event.CHOOSE_DAY):
        choice = datetime.now().strftime('%#d')
        ahk.call('ToolTip', f'Day is {get_choice()}, {HOTKEY_SEND_CHOICE} to insert.')


if __name__ == '__main__':
    main()
