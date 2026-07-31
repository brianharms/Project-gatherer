# Project Gatherer

A portable app that lives on an external drive and gathers every image and video
that has even the slightest chance of belonging to one of your projects — from
any drive on any Mac or Windows machine — into a single, uniquely-named folder on
the drive the app lives on.

- **Cross-platform.** One codebase runs on macOS, Windows, and Linux.
- **Read-only on your data.** Scanned drives are never moved, deleted, or
  modified. Everything is *copied* into a new output folder on the app's drive.
- **High recall by design.** You describe your projects in a Markdown file; the
  app matches those terms against every file's name *and* its full folder path,
  so even a generically-named `IMG_2931.jpg` gets caught if it lives in a
  project folder.
- **Opens PowerPoint.** `.pptx / .pptm / .ppsx / .potx` files are read as zip
  archives; their slide text is matched too, and matching decks have their
  embedded images/videos extracted.

---

## What's on the drive

```
Project-Gatherer/
├─ project_gatherer.py     ← the app (Python, no third-party packages)
├─ Run on Mac.command      ← double-click to launch on a Mac
├─ Run on Windows.bat      ← double-click to launch on Windows
├─ PROJECTS.md             ← REPLACE with your own project descriptions
└─ README.md               ← this file
```

Copy this whole folder to the root of your external drive.

---

## One-time requirement: Python 3

The app is a Python script, so the computer you plug into needs **Python 3.8+**
installed (this keeps the app tiny and easy to inspect — nothing is bundled).
Tkinter, used for the window, ships with the standard installer on both systems.

- **macOS:** Recent versions already include `python3`. If not, install from
  <https://www.python.org/downloads/> (or run `xcode-select --install`).
- **Windows:** Install from <https://www.python.org/downloads/> and tick
  **“Add python.exe to PATH”** during setup.

> Want a true no-install `.app` / `.exe` that bundles Python? See
> [Building standalone apps](#optional-building-standalone-apps) below.

---

## How to use it

1. **Write your projects file.** Open `PROJECTS.md` and replace it with your
   transcribed descriptions of every project you can think of — names, clients,
   places, dates, aliases, codenames, collaborators, folder names you remember.
   The more names and aliases, the higher the recall. (See the tips inside that
   file.) Save it on the drive.
2. **Plug the drive into the Mac or PC** whose drives you want to search.
3. **Launch the app:**
   - macOS → double-click **`Run on Mac.command`**
     (first time, macOS may require: right-click → **Open**).
   - Windows → double-click **`Run on Windows.bat`**.
4. **In the window:**
   - Confirm the projects file (auto-detected, or **Browse…**).
   - Tick the **drives to scan** (internal or USB — click *Refresh* if you plug
     one in after opening).
   - Choose a **sensitivity** (default *High recall* catches almost anything;
     raise it if you get too much noise) and options.
   - Optionally tick **Preview only** first to see how many files match without
     copying anything.
   - Click **Start gathering**.
5. When it finishes, click **Open output folder**.

### What you get

A new folder on the app's drive, named for context, e.g.:

```
ProjectGather__DANA-IMAC__PROJECTS__20260731-142530/
├─ <DriveName>/…                     ← matched files, original folders preserved
│   └─ …/SomeDeck__embedded_media/   ← media pulled out of a PowerPoint
└─ manifest.csv                      ← every file + WHY it matched (its score
                                         and the terms/phrases that triggered it)
```

The `manifest.csv` is the review tool for low-confidence catches: sort by
`score`, read the `reasons` column, and delete anything that clearly doesn't
belong. Nothing on the source drives is affected.

---

## Matching, in short

| Evidence in `PROJECTS.md`                     | Weight    |
|-----------------------------------------------|-----------|
| Markdown heading (`# Project Name`)           | strongest |
| **Bold** or "quoted" text                     | strong    |
| Capitalised proper-noun phrases               | medium    |
| Alphanumeric codes (`PRJ-114`, `aur_v2`)      | medium    |
| Years (`2023`) and generic words              | supporting only |
| Rare distinctive words                        | light     |

A file is matched against **its name and every folder above it**. PowerPoint
decks are also matched on their **slide text**. Everyday words and spoken filler
are ignored so they don't pull in unrelated files.

**Sensitivity** sets the score threshold: *High recall* keeps nearly any real
touch (best for a first sweep), *Balanced* needs a couple of solid hits,
*Strict* keeps only strong matches. **Fuzzy matching** (on by default) also
catches minor typos and variants.

---

## Command-line mode (optional)

The same script runs headless — handy for scripting or machines without a GUI:

```bash
python3 project_gatherer.py --projects PROJECTS.md \
    --drive /Volumes/BackupA --drive /Volumes/BackupB \
    --sensitivity high --scan-only
```

Run `python3 project_gatherer.py --help` for all options. Running it with **no**
arguments launches the graphical app.

---

## Optional: standalone apps (no Python needed)

To run on a machine with **no Python installed**, use a self-contained binary. A
`.app` (Mac) and `.exe` (Windows) built side by side let the same drive work on
either computer. Because PyInstaller can't cross-compile, each binary must be
built on its own OS — so the easiest path is GitHub Actions, which has both.

### Easiest: let GitHub build both for you

A workflow at `.github/workflows/build-apps.yml` builds both binaries on native
runners:

- **Manually:** repo → **Actions** tab → **Build standalone apps** → **Run
  workflow**. When it finishes, download the `ProjectGatherer-macOS` and
  `ProjectGatherer-Windows` artifacts from that run.
- **By tag:** `git tag v1.0 && git push origin v1.0` also attaches zipped
  binaries to a GitHub **Release**.

Unzip both onto your drive and you're done.

### Or build locally

**On a Mac:**
```bash
pip3 install pyinstaller
pyinstaller --onefile --windowed --name "Project Gatherer" project_gatherer.py
# result: dist/Project Gatherer.app
```

**On Windows:**
```bat
pip install pyinstaller
pyinstaller --onefile --windowed --name "Project Gatherer" project_gatherer.py
REM result: dist\Project Gatherer.exe
```

Copy the resulting `.app` / `.exe` (plus your `PROJECTS.md`) onto the drive.

### First-launch security prompt (unsigned builds)

These binaries aren't code-signed (that needs a paid Apple/Microsoft
certificate), so the OS warns once on first launch. This is expected:

- **macOS:** right-click the app → **Open** → **Open**. (If it says the app "is
  damaged", clear the download quarantine once in Terminal:
  `xattr -dr com.apple.quarantine "/Volumes/YourDrive/Project Gatherer.app"`.)
- **Windows:** on the SmartScreen dialog click **More info → Run anyway**.

After the first launch it opens normally. The plain Python script (via the
`Run on…` launchers) doesn't trigger these prompts at all.

---

## Notes & safety

- **Nothing on scanned drives is ever changed** — copy only, never move/delete.
- System, cache, and application-support folders are skipped for speed and to
  avoid false positives; your documents/photos/projects areas are fully scanned.
- Large scans of a whole internal drive can take a while and copy a lot — use
  **Preview only** first, and the `manifest.csv` to prune afterward.
- The app only reads the drives you explicitly tick.
