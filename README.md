# PPTRemote

Control Microsoft PowerPoint from your phone’s browser. No app store install, no Bluetooth. The PC runs a small tray app; the phone joins over the same Wi‑Fi (or hotspot).

## Download

Get **`PPTRemote.exe`** from the latest **[Release](../../releases/latest)**.

One file. Double-click it. Windows does not need a separate .NET install.

The first run may trigger SmartScreen or a firewall prompt. Allow it on private networks so the phone can reach the PC.

## What you need

- Windows 10 or 11
- Microsoft PowerPoint (desktop)
- Phone and PC on the same Wi‑Fi, or the phone on the PC’s hotspot

## Use it

1. Open your deck in PowerPoint.
2. Click the **PPTRemote** tray icon.
3. Scan the QR code, or type the link under it if the camera misses it.
4. Start the slideshow from the phone (or from PowerPoint).

The tray popup shows whether a deck is open or live, which Wi‑Fi you are on, and who is connected.

On the phone:

- **Notes** — speaker notes for the current slide
- **Next slide** — preview the next animation, swipe to advance
- **Prev / Next** — same as clicking through animations
- Slide grid, black/white screen, first/last, exit

Hover the tray icon to see PowerPoint status without opening the popup. Right-click the tray icon → **Quit** to stop the server.

## How it works

The exe hosts a local page on port **8765** (`http://<your-pc-ip>:8765`) and talks to PowerPoint on the same machine. Traffic stays on your LAN. It does not upload the deck.


## License

Use and share freely. No warranty.
