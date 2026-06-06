# 🦇 Stream Bat App

**Stream Bat** is a Windows desktop app that lets you build a single `.bat` (batch) file to launch all your streaming setup at once — OBS, Steam games, regular programs, Microsoft Store apps, and web pages — with one double-click.

> Available in **English** and **Spanish (Español)**.

---

## Downloads

| Version | File |
|---|---|
| English | `Stream_Bat.zip` |
| Spanish | `Stream_Bat_Esp.zip` |

Extract the zip, then run `Stream_Bat.exe` (or `Stream_Bat_Español.exe`). No installation required. 
RECOMMENDED: Leave the Stream Bat app INSIDE the folder so that anything it generates it stays with. If you want to access it from elsewhere, create a shortcut for it and put that shortcut wherever you'd like.

---

## What does Stream Bat do?

When you click **Build Batch File**, Stream Bat:

1. Creates a folder called `Shortcuts_for_bat` next to the `.bat` file, filled with Windows shortcuts (`.lnk` files) for everything you configured.
2. Generates a `.bat` file that loops through those shortcuts and launches them one by one, with a configurable delay between each.

Double-clicking the `.bat` file from then on launches your entire setup automatically.

> ⚠️ **Important:** Do not move or delete the `Shortcuts_for_bat` folder — the batch file depends on it. If you want to move things, rebuild the batch file in the new location.

---

## How to Use Stream Bat

The app is organized into tabs. Work through them from left to right, then hit **Build Batch File**.

---

### OBS Tab

Add your **OBS Studio** executable so it launches with everything else.

- Click **Browse...** to find `obs64.exe` (usually in `C:\Program Files\obs-studio\bin\64bit\`), or paste the path directly into the text box and click **Set Path**.
- The current status is shown below — green means it's set, red means it isn't.
- Use **Clear OBS** to remove it from the launch list.

NOTE: By default, OBS will be the last app to be launched by the batch file you create.

---

### Programs Tab

Add any regular Windows programs (`.exe` files).

1. Enter a **Program Name** (this is just a label for your reference).
2. Either click **Browse...** to locate the `.exe`, or paste the path in directly.
3. Click **Add Program**.

Added programs appear in the list below. Select one and click **Remove Selected** to remove it.

---

### Steam Games/Program Tab

Add Steam games or Programs using their **Steam URL**. Useful for apps like Stream Avatars, VTube Studio, etc.

1. Enter a **Game/Program Name**.
2. Enter the **Steam URL** in the format: `steam://rungameid/XXXXXXXXX`

**How to find a Steam game's URL:**
- Open Steam and go to your Library.
- Right-click the game → **Manage** → **Add desktop shortcut**.
- Right-click the shortcut on your desktop → **Properties**.
- The **Target** field contains the `steam://rungameid/...` URL — copy it.

3. Click **Add Steam Game/Program**.

---

### MS Store Tab

Add apps installed from the **Microsoft Store** (like Spotify, Netflix, etc.).

1. Type the app name in the search box (e.g. `Spotify`) and press **Enter** or click **Search**. NOTE: It is normla behaviour if a Powershell window pops up for a second upon clicking **Search** since Stream Bat is using PowerShell to search for the app's ID.
2. The app automatically queries your system for matching installed Store apps.
3. Select the correct result from the list and click **Add Selected App**.

That's it — no manual ID copying needed.

> If nothing shows up, make sure the app is actually installed from the Microsoft Store (not the desktop/Steam version), and try a shorter search term.

---

### Web Pages Tab

Add URLs to open in your default browser. Useful for text-to-speech websites, YouTube playlists, etc.

1. Enter a **Page Name** (label for your reference).
2. Enter the **URL** (e.g. `https://www.youtube.com`). The `https://` prefix is added automatically if you leave it out.
3. Click **Add Web Page**.

---

### View List Tab

Shows a live summary of everything currently configured across all tabs. Basically a list of all the programs that are "loaded in" before creating the batch file.

---

### Build Tab

Configure and generate the final batch file.

| Setting | Description |
|---|---|
| **Batch File Name** | The name of the `.bat` file that gets created. Defaults to `Open_Stream`. The `.bat` extension is added automatically. |
| **Delay Between Programs** | How many seconds to wait between launching each item (1–30 seconds). Defaults to 2. Increase this if programs are fighting over resources at startup. |

The **Summary** box shows a count of everything that will be included. Once you're ready, click **⚙️ Build Batch File⚙️**.

After a successful build, you'll find:
- `YourFileName.bat` — the launcher (double-click this)
- `Shortcuts_for_bat/` — the folder of shortcuts the batch file uses

---

## 📁 Repo Structure

```
Stream_Bat_Repo/
├── raw_code/
│   ├── Stream_Bat.py               # English source code
│   └── Stream_Bat_Español.py       # Spanish source code
├──Stream_Bat.zip                   # English version of the program
├──Stream_Bat_Esp.zip               # Spanish version of the program
└── README.md
```

---

## 🛠️ Running from Source

**Requirements:**
- Python 3.x
- `pywin32` (`pip install pywin32`)

```bash
pip install pywin32
python raw_code/Stream_Bat.py
```

---

## Compiling to an Executable

The app is compiled using [Auto Py to Exe](https://github.com/brentvollebregt/auto-py-to-exe) (which wraps PyInstaller).

Key settings:
- **Script:** `Stream_Bat.py`
- **One Directory** or **One File** mode
- **Icon:** `Batpic_Icon.ico` (set in the Icon field to embed it in the `.exe` and taskbar)
- **Additional Files:** include `Batpic_Icon.ico` so the app can find it at runtime

---

## Extra Notes

- The app creates real Windows `.lnk` shortcut files, which is why it requires `pywin32`.
- Steam, Microsoft Store, and web shortcuts all use Windows' built-in URL/protocol launching — no third-party dependencies needed beyond `pywin32`.
- OBS is handled slightly differently: it gets its own subfolder inside `Shortcuts_for_bat` so it can be launched with `/wait` (meaning the batch file waits for OBS to open before exiting).

---

## Author

**Timothy Figueroa**
Built with help from Deepseek-V3.
