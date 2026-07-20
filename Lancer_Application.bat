@echo off
echo Verification et installation des dependances...
call npm install
echo Lancement de l'application Master Management...
start http://localhost:3005
call npm run dev
pause