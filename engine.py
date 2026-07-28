"""
Messaging engine — handles WhatsApp sending via pywhatkit with random delays.
"""
import random
import time
import threading
import pywhatkit


class MessageEngine:
    """Sends WhatsApp messages with configurable random delays."""

    def __init__(self, min_delay=20, max_delay=40, wait_time=15, close_time=10, post_send_delay=8):
        self.min_delay = min_delay
        self.max_delay = max_delay
        self.wait_time = wait_time
        self.close_time = close_time
        self.post_send_delay = post_send_delay
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()

    def reset(self):
        self._stop_event.clear()

    def open_whatsapp_web(self):
        """Opens WhatsApp Web in the default browser to check session."""
        import webbrowser
        webbrowser.open("https://web.whatsapp.com")

    def send_message(self, phone: str, message: str) -> tuple[bool, str]:
        """
        Send a single WhatsApp message instantly using pywhatkit.
        Returns (success: bool, detail: str).
        """
        # Auto-normalize: add '+' if missing
        phone = phone.strip()
        if not phone.startswith('+'):
            if phone.startswith('00'):
                phone = '+' + phone[2:]
            else:
                # If no '+' and no '00', we'll assume it needs a '+' 
                # (and maybe a country code, but main.py should handle that)
                phone = '+' + phone

        try:
            # pywhatkit handles opening/typing. We keep the tab open and do the
            # final Enter/close ourselves so WhatsApp has enough time to load.
            pywhatkit.sendwhatmsg_instantly(
                phone_no=phone,
                message=message,
                wait_time=self.wait_time,       # seconds to wait for WhatsApp Web to load
                tab_close=False,
                close_time=self.close_time,
            )
            self._confirm_send_and_close_tab()
            return True, "Sent"
        except Exception as exc:
            return False, str(exc)

    def _confirm_send_and_close_tab(self):
        """Press Enter to send, then close the active WhatsApp tab."""
        try:
            import pyautogui
        except Exception:
            return

        # Give WhatsApp Web extra time after pywhatkit returns. On slower
        # connections the message can be typed before the send button is ready.
        time.sleep(self.post_send_delay)
        pyautogui.press("enter")
        # Wait after Enter so the message request is actually sent before close.
        time.sleep(self.close_time)

        # Close only the active WhatsApp tab. Retrying can close another tab.
        pyautogui.hotkey("ctrl", "w")

    def send_batch(self, recipients: list[dict], template: str,
                   on_progress=None, on_complete=None, on_delay=None):
        """
        Send messages to a batch of recipients on a background thread.

        `recipients` — list of dicts with keys: name, phone, grade, row_index
        `template` — message string with {name} / {grade} placeholders
        `on_progress(row_index, success, detail)` — called after each send
        `on_complete()` — called when batch finishes
        `on_delay(seconds_remaining, total_delay)` — called every second during delay
        """
        self.reset()

        def _worker():
            for i, r in enumerate(recipients):
                if self._stop_event.is_set():
                    break

                full_name = str(r.get("name", "")).strip()
                first_name = full_name.split()[0] if full_name else ""
                msg = template.replace("{name}", first_name).replace("{group_code}", r.get("group_code", "")).replace("{time}", r.get("time", ""))
                ok, detail = self.send_message(r["phone"], msg)

                if on_progress:
                    on_progress(r["row_index"], ok, detail)

                # Random delay between messages (skip after the last one)
                if i < len(recipients) - 1 and not self._stop_event.is_set():
                    delay = random.randint(self.min_delay, self.max_delay)
                    for sec in range(delay):
                        if self._stop_event.is_set():
                            break
                        remaining = delay - sec
                        if on_delay:
                            on_delay(remaining, delay)
                        time.sleep(1)
                    # Signal delay ended (0 remaining)
                    if on_delay and not self._stop_event.is_set():
                        on_delay(0, delay)

            if on_complete:
                on_complete()

        t = threading.Thread(target=_worker, daemon=True)
        t.start()
        return t
