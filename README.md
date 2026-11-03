<p align="center"><img src="icon/AppIcon.png" width="128" alt=""></p>

# Prism Extractor

A small Mac app that pulls embedded **GraphPad Prism graphs out of PowerPoint decks**, so you can
open and edit them in Prism again.

Office for Mac can't edit embedded Prism objects (Windows can). If you've been handed a deck full
of Prism graphs and no original files, drop the deck on this app and you get the graphs back as
Prism files.

It's built on [prism-ole-handler](https://github.com/armish/prism-ole-handler) by B. Arman Aksoy.

## Download

1. Download **Prism-Extractor-macOS.zip** from the
   [latest release](https://github.com/markusviks/prism-extractor/releases/latest) and unzip it.
2. Drag **Prism Extractor** into your Applications folder.
3. The first time you open it, macOS blocks it, because the app isn't signed by a paid Apple
   developer account. Click **Done**, then go to **System Settings → Privacy & Security**, scroll
   down and click **Open Anyway** next to "Prism Extractor". You only need to do this once.

Works on Apple silicon and Intel Macs.

## Use

- Drop one or more `.pptx` files on the window, or click to choose them.
- Each graph is listed with its slide number. Then:
  - **Open in Prism** saves the graph to your Downloads folder and opens it.
  - **Download** saves it to Downloads and shows it in Finder.
  - **Download all** zips every graph into one file in Downloads, named after the deck.
- Clicking again won't make duplicates. If you've edited a saved graph in Prism, the next save
  goes to a new name such as `… (1).pzf`, so your edits are never overwritten.
- Your files never leave your Mac. The app runs entirely on your computer.

### Which Prism versions?

| Graph made with | You get | Notes |
| --- | --- | --- |
| Prism 10+ | `.prism` | |
| Prism 4–9 | `.pzf` | `prism-ole-handler` skips these; this app unwraps them itself. Tested on Prism 8 graphs. |
| Prism 3 and earlier | `.pzm` | Untested. |

Graphs from Word (`.docx`) and Excel (`.xlsx`) files work too, but that's experimental. Old
binary `.ppt` files aren't supported; open them in PowerPoint and save as `.pptx` first.

Prism 10 graphs come out of `prism-ole-handler` with a `.pzfx` name even though they are ZIP-based
`.prism` files. This app names every file by what it actually contains.

## Run from source

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py            # app window
.venv/bin/python app.py --browser  # or in your web browser at http://127.0.0.1:8765
```

Tests: `.venv/bin/pip install pytest && .venv/bin/python -m pytest`

## Build the app

```bash
./build_app.sh            # universal2 (needs the python.org Python); or: ./build_app.sh arm64
```

The app ends up in `dist/`. Pushing a tag like `v1.0.0` builds it on GitHub Actions and attaches
the zip to a release.

The icon is drawn by `icon/make_icon.py` (needs Pillow), which also rebuilds `icon/AppIcon.icns`.

## License

MIT. See [LICENSE](LICENSE).
