import asyncio
import functools
import inspect
import os
import re
import sys
import time
import tkinter
import tkinter.constants
import tkinter.scrolledtext
import tkinter.ttk
import tkinter.filedialog
import json
from excel_processor import ExcelProcessor
from html_formatter import prepare_body
from telethon import TelegramClient, events, utils
import imgkit
from datetime import datetime
import os
img_opts = {
  "enable-local-file-access": None
}
with open('style.css') as f:
    css=f.read
IMAGE = os.path.dirname(os.path.abspath(__file__)) + '/name.jpg'
# Some configuration for the app
TITLE = 'Report distributor'
SIZE = '640x280'
REPLY = re.compile(r'\.r\s*(\d+)\s*(.+)', re.IGNORECASE)
DELETE = re.compile(r'\.d\s*(\d+)', re.IGNORECASE)
EDIT = re.compile(r'\.s(.+?[^\\])/(.*)', re.IGNORECASE)
TEMPLATE_START_LOG = '{} - Sending started phone: {}, contract number: {}, records: {}, image count: {}\n'
TEMPLATE_END_LOG = '{} - Sending end.\n'
def get_env(name, message, cast=str):
    if name in os.environ:
        return os.environ[name]
    while True:
        value = input(message)
        try:
            return cast(value)
        except ValueError as e:
            print(e, file=sys.stderr)
            time.sleep(1)


with open('config.json', 'r') as f:
    config = json.load(f)

# Session name, API ID and hash to use; loaded from environmental variables
SESSION = os.environ.get('TG_SESSION', 'gui')
API_ID = config['api_id']
API_HASH = config['api_hash']


def sanitize_str(string):
    return ''.join(x if ord(x) <= 0xffff else
                   '{{{:x}ū}}'.format(ord(x)) for x in string)

def callback(func):
    @functools.wraps(func)
    def wrapped(*args, **kwargs):
        result = func(*args, **kwargs)
        if inspect.iscoroutine(result):
            aio_loop.create_task(result)

    return wrapped


def allow_copy(widget):
    widget.bind('<Control-c>', lambda e: None)
    widget.bind('<Key>', lambda e: "break")
class App(tkinter.Tk):
    def __init__(self, client, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.cl = client
        self.me = None

        self.title(TITLE)
        self.geometry(SIZE)

        # Signing in row; the entry supports phone and bot token
        self.sign_in_label = tkinter.Label(self, text='Loading...')
        self.sign_in_label.grid(row=0, column=0)
        self.sign_in_entry = tkinter.Entry(self)
        self.sign_in_entry.grid(row=0, column=1, sticky=tkinter.EW)
        self.sign_in_entry.bind('<Return>', self.sign_in)
        self.sign_in_button = tkinter.Button(self, text='...',
                                             command=self.sign_in)
        self.sign_in_button.grid(row=0, column=2)
        self.code = None

        # The chat where to send and show messages from
        tkinter.Label(self, text='Target:').grid(row=1, column=0)

        self.chat = tkinter.Entry(self)
        self.chat.grid(row=1, column=1, sticky=tkinter.EW)
        self.columnconfigure(1, weight=1)
        tkinter.Button(self, text='Browse',
                       command=self.choose_file).grid(row=1, column=2)
        # Message log (incoming and outgoing); we configure it as readonly
        self.log = tkinter.scrolledtext.ScrolledText(self)
        allow_copy(self.log)
        self.log.grid(row=2, column=0, columnspan=3, sticky=tkinter.NSEW)
        self.rowconfigure(2, weight=1)
        self.cl.add_event_handler(self.on_message, events.NewMessage)

        # Save shown message IDs to support replying with ".rN reply"
        # For instance to reply to the last message ".r1 this is a reply"
        # Deletion also works with ".dN".
        self.message_ids = []

        self.send_message_btn = tkinter.Button(self, text='Send',
                       command=self.send_message)
        self.send_message_btn.grid(row=3, column=2)

        # Post-init (async, connect client)
        self.cl.loop.create_task(self.post_init())
    async def post_init(self):
        if await self.cl.is_user_authorized():
            self.set_signed_in(await self.cl.get_me())
        else:
            # User is not logged in, configure the button to ask them to login
            self.sign_in_button.configure(text='Sign in')
            self.sign_in_label.configure(
                text='Sign in (phone):')

    async def on_message(self, event):
        """
        Event handler that will add new messages to the message log.
        """
        # We want to show only messages sent to this chat
        if event.chat_id != self.chat_id:
            return

        # Save the message ID so we know which to reply to
        self.message_ids.append(event.id)

        # Decide a prefix (">> " for our messages, "<user>" otherwise)
        if event.out:
            text = '>> '
        else:
            sender = await event.get_sender()
            text = '<{}> '.format(sanitize_str(
                utils.get_display_name(sender)))

        # If the message has media show "(MediaType) "
        if event.media:
            text += '({}) '.format(event.media.__class__.__name__)

        text += sanitize_str(event.text)
        text += '\n'

        # Append the text to the end with a newline, and scroll to the end
        self.log.insert(tkinter.END, text)
        self.log.yview(tkinter.END)

    # noinspection PyUnusedLocal
    @callback
    async def sign_in(self, event=None):
        """
        Note the `event` argument. This is required since this callback
        may be called from a ``widget.bind`` (such as ``'<Return>'``),
        which sends information about the event we don't care about.
        This callback logs out if authorized, signs in if a code was
        sent or a bot token is input, or sends the code otherwise.
        """
        if await self.cl.is_user_authorized():
            await self.cl.log_out()
            self.destroy()
            return
        value = self.sign_in_entry.get().strip()
        if not value:
            return

        self.sign_in_label.configure(text='Working...')
        self.sign_in_entry.configure(state=tkinter.DISABLED)
        if self.code:
            self.set_signed_in(await self.cl.sign_in(code=value))
        elif ':' in value:
            self.set_signed_in(await self.cl.sign_in(bot_token=value))
        else:
            self.code = await self.cl.send_code_request(value)
            self.sign_in_label.configure(text='Code:')
            self.sign_in_entry.configure(state=tkinter.NORMAL)
            self.sign_in_entry.delete(0, tkinter.END)
            self.sign_in_entry.focus()
            return

    def set_signed_in(self, me):
        """
        Configures the application as "signed in" (displays user's
        name and disables the entry to input phone/bot token/code).
        """
        self.me = me
        self.sign_in_label.configure(text='Signed in')
        self.sign_in_entry.configure(state=tkinter.NORMAL)
        self.sign_in_entry.delete(0, tkinter.END)
        self.sign_in_entry.insert(tkinter.INSERT, utils.get_display_name(me))
        self.sign_in_entry.configure(state=tkinter.DISABLED)
        self.sign_in_button.configure(text='Log out')
        self.chat.focus()

    # noinspection PyUnusedLocal
    @callback
    async def send_message(self, event=None):
        if not self.cl.is_connected():
            return
        self.chat.configure(state=tkinter.DISABLED)
        self.send_message_btn.configure(state=tkinter.DISABLED)
        excel_processor = ExcelProcessor(self.chat.get())
        result = excel_processor.process_file()
        print('Excel content is parsed. Sending started...')
        for key in list(result.keys()):
            text = TEMPLATE_START_LOG.format(datetime.now().strftime("%m/%d/%Y, %H:%M:%S") , 
                result[key].phone ,
                result[key].contract, 
                result[key].count,
                len(result[key].contents ))
            self.log.insert(tkinter.END, text)
            self.log.yview(tkinter.END)
            user = await self.cl.get_entity(result[key].phone)
            await self.cl.send_message(user, str(result[key].contract))
            for content in result[key].contents:
                html = prepare_body(excel_processor.header+content)
                imgkit.from_string(html, IMAGE, options=img_opts )
                await self.cl.send_file(user, IMAGE)
            text = TEMPLATE_END_LOG.format( datetime.now().strftime("%m/%d/%Y, %H:%M:%S"))
            self.log.insert(tkinter.END, text)
        self.chat.configure(state=tkinter.NORMAL)
        self.send_message_btn.configure(state=tkinter.NORMAL)

    @callback
    async def choose_file(self, event=None):
        if not self.cl.is_connected():
            return
        self.file_name = tkinter.filedialog.askopenfile(
            mode='r', title='Choose a file').name
        self.chat.delete(0, tkinter.END)
        self.chat.configure(state=tkinter.NORMAL)
        self.chat.insert(tkinter.INSERT, self.file_name)
        self.chat.focus()


async def main(interval=0.05):
    client = TelegramClient(SESSION, API_ID, API_HASH)
    try:
        await client.connect()
    except Exception as e:
        print('Failed to connect', e, file=sys.stderr)
        return

    app = App(client)
    try:
        while True:
            app.update()
            await asyncio.sleep(interval)
    except KeyboardInterrupt:
        pass
    except tkinter.TclError as e:
        if 'application has been destroyed' not in e.args[0]:
            raise
    finally:
        await app.cl.disconnect()


if __name__ == "__main__":
    # Some boilerplate code to set up the main method
    aio_loop = asyncio.get_event_loop()
    try:
        aio_loop.run_until_complete(main())
    finally:
        if not aio_loop.is_closed():
            aio_loop.close()
