import time
import json
from pynput import mouse, keyboard

class Recorder:
    def __init__(self):
        self.events = []
        self.start_time = time.time()

    def on_move(self, x, y):
        self.events.append({'time': time.time() - self.start_time, 'type': 'move', 'x': x, 'y': y})

    def on_click(self, x, y, button, pressed):
        self.events.append({
            'time': time.time() - self.start_time,
            'type': 'click',
            'x': x,
            'y': y,
            'button': str(button),
            'pressed': pressed
        })

    def on_scroll(self, x, y, dx, dy):
        self.events.append({
            'time': time.time() - self.start_time,
            'type': 'scroll',
            'x': x,
            'y': y,
            'dx': dx,
            'dy': dy
        })

    def on_press(self, key):
        try:
            self.events.append({'time': time.time() - self.start_time, 'type': 'press', 'key': key.char})
        except AttributeError:
            self.events.append({'time': time.time() - self.start_time, 'type': 'press', 'key': str(key)})

    def on_release(self, key):
        try:
            self.events.append({'time': time.time() - self.start_time, 'type': 'release', 'key': key.char})
        except AttributeError:
            self.events.append({'time': time.time() - self.start_time, 'type': 'release', 'key': str(key)})
        if key == keyboard.Key.esc:
            return False

    def start_recording(self):
        self.start_time = time.time()
        with mouse.Listener(on_move=self.on_move, on_click=self.on_click, on_scroll=self.on_scroll) as mouse_listener, \
             keyboard.Listener(on_press=self.on_press, on_release=self.on_release) as keyboard_listener:
            mouse_listener.join()
            keyboard_listener.join()

    def save(self, filename):
        with open(filename, 'w') as f:
            json.dump(self.events, f, indent=4)

class Player:
    def __init__(self, filename):
        with open(filename, 'r') as f:
            self.events = json.load(f)

    def play(self):
        mouse_controller = mouse.Controller()
        keyboard_controller = keyboard.Controller()

        start_time = time.time()
        for event in self.events:
            time.sleep(event['time'] - (time.time() - start_time))
            if event['type'] == 'move':
                mouse_controller.position = (event['x'], event['y'])
            elif event['type'] == 'click':
                button = mouse.Button.left if 'left' in event['button'] else mouse.Button.right
                if event['pressed']:
                    mouse_controller.press(button)
                else:
                    mouse_controller.release(button)
            elif event['type'] == 'scroll':
                mouse_controller.scroll(event['dx'], event['dy'])
            elif event['type'] == 'press':
                key = event['key']
                if len(key) > 1 and key.startswith('Key.'):
                    key = getattr(keyboard.Key, key.split('.')[1])
                keyboard_controller.press(key)
            elif event['type'] == 'release':
                key = event['key']
                if len(key) > 1 and key.startswith('Key.'):
                    key = getattr(keyboard.Key, key.split('.')[1])
                keyboard_controller.release(key)

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description="Record and play back mouse and keyboard events.")
    parser.add_argument('mode', choices=['record', 'play'], help="Mode: 'record' to record events, 'play' to play events")
    parser.add_argument('file', help="File to save to or play from")
    args = parser.parse_args()

    if args.mode == 'record':
        print("Recording... Press ESC to stop.")
        recorder = Recorder()
        recorder.start_recording()
        recorder.save(args.file)
        print(f"Recording saved to {args.file}")
    elif args.mode == 'play':
        print(f"Playing back events from {args.file}...")
        player = Player(args.file)
        player.play()
        print("Playback finished.")
