import win32clipboard
import win32con
import logging
import base64

def validate_base64(data):
    """Validate if a string is valid base64 data."""
    try:
        base64.b64decode(data)
        return True
    except Exception:
        return False

def set_clipboard_html(html):
    """Set HTML content to the clipboard."""
    try:
        CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
        html_header = (
            "Version:0.9\r\n"
            "StartHTML:00000000\r\n"
            "EndHTML:00000000\r\n"
            "StartFragment:00000000\r\n"
            "EndFragment:00000000\r\n"
        )
        html_content = f"<!DOCTYPE html><html><body><!--StartFragment-->{html}<!--EndFragment--></body></html>"
        start_html = len(html_header)
        start_fragment = start_html + html_content.find("<!--StartFragment-->") + len("<!--StartFragment-->")
        end_fragment = start_html + html_content.find("<!--EndFragment-->")
        end_html = start_html + len(html_content)
        html_header = (
            f"Version:0.9\r\n"
            f"StartHTML:{start_html:08d}\r\n"
            f"EndHTML:{end_html:08d}\r\n"
            f"StartFragment:{start_fragment:08d}\r\n"
            f"EndFragment:{end_fragment:08d}\r\n"
        )
        final_html = html_header + html_content
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(CF_HTML, final_html.encode('utf-8'))
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        logging.error(f"Error setting HTML clipboard: {e}")
        return False

def get_clipboard_text():
    """Retrieve text from the clipboard."""
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                if text:
                    logging.info(f"Retrieved clipboard text: {text[:100]}...")
                return text
            else:
                logging.info("Clipboard contains non-text data")
                return None
        finally:
            win32clipboard.CloseClipboard()
    except win32clipboard.error as e:
        logging.error(f"Error accessing clipboard: {e}")
        return None