# Sığorta Hesabatı — Quraşdırma Təlimatı

> Send this file to your colleagues along with the `.zip` for their platform.

## Windows

1. Download `Sigorta-Hesabati-Windows.zip`.
2. Right-click the zip → **Extract All…** → pick a location (Desktop is fine). A folder named **Sigorta Hesabati** appears.
3. Open that folder and double-click **`Sigorta Hesabati.exe`**.
4. Windows will probably show a blue **"Windows protected your PC"** popup the first time:
   - Click **More info**
   - Click **Run anyway**
   This warning only appears the first time. (It pops up because the program is unsigned — that's normal for in-house software; it is not a virus.)
5. A small black console window opens, and shortly after your browser opens to the app. Done.
6. **To quit:** close the black console window.

## macOS

1. Download `Sigorta-Hesabati-macOS.zip` and double-click it to extract. You'll get **`Sigorta Hesabati.app`**.
2. **First launch — important:** **right-click** (or Ctrl-click) the app → **Open** → in the popup that says *"Apple could not verify..."*, click **Open**.
   - If you just double-click, macOS blocks it with no easy way to bypass.
   - This dance is only needed the very first time. After that, double-clicking works normally.
3. Your browser opens automatically. Done.
4. **To quit:** in the menu bar at the top, click **Sigorta Hesabati** → **Quit Sigorta Hesabati** (or press ⌘Q).

If macOS *still* blocks it after right-click → Open, go to **System Settings → Privacy & Security**, scroll down, you'll see *"Sigorta Hesabati was blocked..."* — click **Open Anyway**.

## Using the app

1. Click the upload area (or drag the Excel file onto it).
2. Pick the `.xlsx` file. Click **Faylı işlə**. Wait a few seconds — for ~60 000-row files this takes 10–20 seconds.
3. You'll see a list of all insurance companies. From here you can:
   - Click an insurance name to preview its rows
   - Click **Yüklə** to download just that insurance's `.xlsx`
   - Click **Hamısını yüklə (.zip)** at the top to download all insurances as a ZIP archive of separate `.xlsx` files

## Privacy

The app runs entirely on your computer. No data is sent over the internet. The uploaded file lives only in your computer's memory and is wiped when you quit the app.

## Troubleshooting

| Problem | Fix |
|---|---|
| Browser didn't open | Open any browser, go to <http://127.0.0.1:5050> |
| Browser shows "site can't be reached" | The app hasn't started yet — wait 5 seconds and refresh |
| "Boş port tapılmadı" error | Another program is using ports 5050–5099. Quit Skype/other apps and try again |
| (macOS) "damaged and can't be opened" | Run in Terminal: `xattr -dr com.apple.quarantine "/path/to/Sigorta Hesabati.app"` |
