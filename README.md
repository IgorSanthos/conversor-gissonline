echo "# gerenciador_desktop" >> README.md
git init
git add README.md
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/IgorSanthos/gerenciador_desktop.git
git push -u origin main


pyinstaller --onefile -w app.py

deactivate
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install pyinstaller