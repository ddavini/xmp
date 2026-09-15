 # X-MaD.Player (XmP)

X-MaD.Player, known as "XmP," is a Windows MP3 player written in Visual Basic 6, originally developed between roughly 2001 and 2005 (this snapshot is version 1.0.378, code name "Gamera"). It was recovered from a personal CVS backup and is published here as-is, for archival purposes.

## What's in here

- **`Source/`** - the main XmP application: the player UI, EQ, visualizations (FFT spectrum, oscilloscope), hotkeys, playlist handling, ID3 tag editing, and the `xmINI Editor` / `xmSCROLLER` companion tools bundled alongside it.
- **`Source/VC++/`** - native C++ engines the player is built on: the custom `xmMP3`/`VBMP3` decompression engine (itself based on GPL'd `mp3dec` code credited to Xing Technology/GoodNoise), the retired `xmFFT` spectrum engine, and an unused `xmOGG` experiment.
- **`Source Other/`** - separate, related companion projects that talk to the player at runtime rather than being compiled into it: `MicroHandler` (a small bundled web server used for a remote LCD display feature) and `LCD232` (the corresponding LCD driver/client), plus `vbConsoleLinker`, a VB6-era build-time tool.

Everything unrelated to the actual project (unused bundled third-party SDKs, an unrelated audio-engine demo, an unrelated crypto library) was removed before publishing; see the commit history for details.

## License

GPLv2 - see [`LICENSE`](LICENSE). The original `readme.txt`/`credits.txt` declared the project GPL-licensed, and its core MP3 engine is itself a GPL derivative, so the license carries through to the whole codebase.

## A note on AI involvement

All of the code in this repository was written by human authors, years before this repository existed. Claude (Anthropic) was used only to help prepare this old codebase for publishing here, specifically:

- **Proofreading**: fixing English spelling/grammar typos in comments, documentation, UI strings, and error messages.
- **Machine translation**: translating scattered Italian comments and two bundled readmes into English, added *alongside* the original Italian text (nothing was deleted or replaced), so both languages are still visible.
- **Repository cleanup**: identifying and removing bundled third-party code that was never actually used by the project, and setting up `.gitignore`/git for the initial publish.

Claude did not write, generate, or modify any application logic, and no source-code identifiers (module, function, or variable names - many of which are still in Italian) were touched.
