# Dragon Bridge for LibreOffice

Bridges Dragon NaturallySpeaking with LibreOffice Writer, enabling direct speech-to-text dictation and voice commands.

## Features

### Clipboard Bridge (Auto-Transfer)
Monitors the Windows clipboard for changes. When Dragon outputs text, it automatically appears in your LibreOffice Writer document at the cursor position — no manual copy/paste needed.

**How to use:**
1. Open a document in LibreOffice Writer
2. Go to **Tools > Dragon Bridge > Start/Stop Clipboard Bridge** (or use the toolbar button)
3. Start dictating with Dragon — text flows directly into your document

### Voice Command Processing
Recognizes common Dragon voice commands that come through as text when Full Text Control isn't available:

- **Punctuation:** "period", "comma", "question mark", "exclamation point", "colon", "semicolon", "open quote", "close quote"
- **Navigation:** "new line", "new paragraph", "tab key"
- **Editing:** "undo that", "scratch that", "redo that", "select all", "delete that"
- **Formatting:** "bold that", "italicize that", "underline that"
- **Clipboard:** "cut that", "copy that", "paste that"
- **File:** "save document"

## Settings

Go to **Tools > Dragon Bridge > Settings** to configure:
- **Clipboard poll interval** (default 300ms) — how often to check for clipboard changes
- **Auto-space** — automatically add a space before inserted text
- **Status bar notifications** — show bridge status in the status bar

## Installation

1. Double-click DragonBridge.oxt or go to **Tools > Extension Manager > Add**
2. Select DragonBridge.oxt
3. Restart LibreOffice

## Requirements

- Windows 10/11
- LibreOffice 7.0+ (Writer)
- Dragon NaturallySpeaking 15/16 or Dragon Professional

## Tips

- If clipboard monitoring seems slow, try reducing the poll interval in Settings
- The extension only monitors when you explicitly turn it on — it won't interfere with normal clipboard use otherwise

## License

MIT License — Copyright (c) 2025 Sami Ahmed
