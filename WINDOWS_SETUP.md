# Windows verzia

## 1) Spustenie bez buildovania `.exe`

Najjednoduchsie je spustit:

`run_gui_windows.bat`

Skript:
- vytvori `.venv` (ak este nie je),
- nainstaluje baliky z `requirements.txt`,
- spusti `desktop_app.py`.

## 2) Build `.exe` (PyInstaller)

Na Windows spusti v PowerShell:

```powershell
.\build_windows_exe.ps1
```

Vysledok:

`dist\TrialBalanceApp\TrialBalanceApp.exe`

## 3) Co dat ku aplikacii

Pri distribucii nechaj spolu:
- `TrialBalanceApp.exe`
- `config.yaml`
- `templates\Trial Balance_template.xlsx`
- priecinok `input\` (pre vstupy)

`output\` sa vytvori automaticky pri behu.
